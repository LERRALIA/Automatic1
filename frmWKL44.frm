VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL44 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Favoritenliste"
   ClientHeight    =   8625
   ClientLeft      =   3105
   ClientTop       =   2055
   ClientWidth     =   11910
   Icon            =   "frmWKL44.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   11
      Left            =   11400
      TabIndex        =   24
      Top             =   120
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Favoriten-Liste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   11655
      Begin VB.CheckBox Check2 
         Caption         =   "nur mit Bestand"
         Height          =   255
         Left            =   5400
         TabIndex        =   27
         Top             =   3600
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "nur geführte Artikel"
         Height          =   255
         Left            =   5400
         TabIndex        =   26
         Top             =   3360
         Width           =   3135
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   25
         Top             =   1320
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
         Caption         =   "Diverse"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   855
         Left            =   5400
         TabIndex        =   21
         Top             =   2160
         Width           =   4095
         Begin VB.OptionButton Option2 
            Caption         =   "aufsteigend"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Tag             =   "asc"
            Top             =   600
            Width           =   2295
         End
         Begin VB.OptionButton Option2 
            Caption         =   "absteigend"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Tag             =   "desc"
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin sevCommand3.Command Command0 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   3720
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List3 
         Height          =   1425
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   19
         Top             =   2160
         Width           =   3975
      End
      Begin sevCommand3.Command Command0 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1800
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
      Begin sevCommand3.Command Command0 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   120
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   2
         Left            =   120
         MaxLength       =   20
         TabIndex        =   14
         Top             =   480
         Width           =   2895
      End
      Begin sevCommand3.Command Command0 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   960
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1320
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Verkaufsmenge Vorjahr"
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
         Index           =   2
         Left            =   5400
         TabIndex        =   10
         Tag             =   "VKMENGE30"
         Top             =   960
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Verkaufsmenge aktueller Monat"
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
         Left            =   5400
         TabIndex        =   8
         Tag             =   "VKMENGE40"
         Top             =   1320
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Verkaufsmenge Vormonat"
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
         Left            =   5400
         TabIndex        =   7
         Tag             =   "VKMENGE50"
         Top             =   1680
         Width           =   3855
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   3
         Top             =   720
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   2
         Top             =   120
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
         Caption         =   "Suche Daten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Verkaufsmenge aktuelles Jahr"
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
         Index           =   1
         Left            =   5400
         TabIndex        =   1
         Tag             =   "VKMENGE20"
         Top             =   600
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bestand"
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
         Index           =   0
         Left            =   5400
         TabIndex        =   0
         Tag             =   "Bestand"
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Linie"
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
         Index           =   1
         Left            =   600
         TabIndex        =   18
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferant"
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
         Left            =   600
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Marke"
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
         Index           =   9
         Left            =   600
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
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
         Left            =   5400
         TabIndex        =   9
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label lblAnzeige 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   8040
      Width           =   10815
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Favoritenliste"
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
      TabIndex        =   5
      Top             =   0
      Width           =   10935
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
End
Attribute VB_Name = "frmWKL44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
    Case 11
        gsHelpstring = "Favoritenliste"
        frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Warenverteilung Verwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    voreinstellungspeichern44
    loeschNEW "VKFAV", gdBase
    loeschNEW "FAVKOPF", gdBase
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
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0
            LeseUmsatzDaten
        Case Is = 1
            Unload frmWKL44
        Case Is = 2
            frmWKL103.Show 1
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function fnPruefeEingabeWKL10()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim sSQL As String
    
    fnPruefeEingabeWKL10 = 1
    
    If Trim$(Text1(7).Text) <> "" Then
    
        If LoeseMarkenstringinLPZ12(Trim$(Text1(7).Text)) = True Then
            fnPruefeEingabeWKL10 = 0
            Exit Function
        Else
            Text1(7).Text = ""
        End If
    Else
        sSQL = "Delete from  MA" & srechnertab
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Trim$(Text1(2).Text) <> "" Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If
   
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWKL10"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Private Sub LeseUmsatzDaten()
    On Error GoTo LOKAL_ERROR
    
    Dim cVon        As String
    Dim cBis        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim rsrs        As Recordset
    Dim cSQL        As String
    Dim cLinr       As String
    Dim clpz        As String
    Dim lcount      As Long
    Dim cMarke      As String
    Dim bAnd        As Boolean
    Dim iRet        As Integer
    Dim corder      As String
    Dim bVKM        As Boolean
    
    bVKM = False
    bAnd = False
    
'    iRet = fnPruefeEingabeWKL10()
'    If iRet <> 0 Then
'        anzeige "rot", "Bitte mindestens ein Suchkriterium angeben!", lblAnzeige
'        Text1(2).SetFocus
'        Exit Sub
'    End If

    anzeige "normal", "Daten werden ermittelt...", lblAnzeige
    
    cLinr = ""
    If Text1(2).Text <> "" Then
        If IsNumeric(Text1(2).Text) Then
            cLinr = Text1(2).Text
        End If
    End If
    
    cMarke = ""
    If Text1(7).Text <> "" Then
        cMarke = Text1(7).Text
    End If
    

    loeschNEW "VKFAV", gdBase
    CreateTable "VKFAV", gdBase
    
    cSQL = "Insert into VKFAV "
    cSQL = cSQL & "Select "
    cSQL = cSQL & "ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", BESTAND "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", aufdat "
    cSQL = cSQL & ", exdat "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & ", Val(awm) as farbnr "
    cSQL = cSQL & " from ARTIKEL  "
    
    If cMarke <> "" Then
        If Datendrin("MA" & srechnertab, gdBase) Then
            If bAnd Then
                cSQL = cSQL & " and "
            Else
                cSQL = cSQL & " where "
            End If
            cSQL = cSQL & " artnr in (Select artnr from MA" & srechnertab & ") "
            bAnd = True
        End If
    End If
    
    If cLinr <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        
        cSQL = cSQL & " LINR = " & cLinr
        bAnd = True
    End If
    
    If Check1.Value = vbChecked Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        
        cSQL = cSQL & " gefuehrt = 'J' "
        bAnd = True
    End If
    
    If List3.ListCount <> 0 Then
    
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        
        cSQL = cSQL & "  (LPZ = " & Mid(List3.list(0), 1, InStr(1, List3.list(0), " ")) & " "
        For lcount = 1 To List3.ListCount - 1
            cSQL = cSQL & " or LPZ = " & Mid(List3.list(lcount), 1, InStr(1, List3.list(lcount), " ")) & " "
        Next lcount
        cSQL = cSQL & ")"
    End If
    
    If Check2.Value = vbChecked Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        
        cSQL = cSQL & " BESTAND > 0 "
        bAnd = True
    
    
    End If
    
    
    
    
    
    
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from VKFAV "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            anzeige "normal", lcount & " " & rsrs!BEZEICH, lblAnzeige
            lcount = lcount - 1
    
            rsrs!VKMENGE20 = vklj(rsrs!artnr, "Anzahl")
            rsrs!VKMENGE30 = vkvj(rsrs!artnr, "Anzahl")
            rsrs!VKMENGE40 = vkam(rsrs!artnr, "Anzahl")
            rsrs!VKMENGE50 = vkvm(rsrs!artnr, "Anzahl")
    
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

    loeschNEW "FAVKOPF", gdBase
    CreateTable "FAVKOPF", gdBase
    
    cSQL = "Select * from FAVKOPF"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    rsrs.AddNew
    If cLinr <> "" Then
        rsrs!linr = cLinr
        rsrs!LIEFBEZ = ermLiefBez(CLng(cLinr))
    End If
    
    
    rsrs!MARKE = cMarke
    If List3.ListCount > 0 Then
        If List3.ListCount > 1 Then
            rsrs!linbez = "verschiedene Linien"
        Else
            rsrs!linbez = List3.list(0)
        End If
    
    End If
    
    If Option1(0).Value Then
        rsrs!sOrt = Option1(0).Caption
    ElseIf Option1(1).Value Then
        rsrs!sOrt = Option1(1).Caption
    ElseIf Option1(2).Value Then
        rsrs!sOrt = Option1(2).Caption
    ElseIf Option1(3).Value Then
        rsrs!sOrt = Option1(3).Caption
    ElseIf Option1(4).Value Then
        rsrs!sOrt = Option1(4).Caption
    End If
    
    If Option2(0).Value Then
         rsrs!sOrt = rsrs!sOrt & " " & Option2(0).Caption
    ElseIf Option2(1).Value Then
         rsrs!sOrt = rsrs!sOrt & " " & Option2(1).Caption
    End If
    
    If Check1.Value = vbChecked Then
        rsrs!sgef = "nur geführte Artikel"
    End If
    
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    BringFarbeInsSpiel "VKFAV", gdBase
    
    cSQL = "Update VKFAV inner join LINBEZ on VKFAV.linr = LINBEZ.linr and VKFAV.lpz = LINBEZ.lpz "
    cSQL = cSQL & " Set VKFAV.marke = LINBEZ.Marke "
    cSQL = cSQL & " , VKFAV.LINBEZ = LINBEZ.LINBEZEICH"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "FAVT", gdBase
    cSQL = "Select * into FAVT from VKFAV"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "VKFAV", gdBase
    CreateTable "VKFAV", gdBase
    
    
    If Option1(0).Value Then
        corder = " order by " & Option1(0).Tag

    ElseIf Option1(1).Value Then
        corder = " order by " & Option1(1).Tag
        bVKM = True
    ElseIf Option1(2).Value Then
        corder = " order by " & Option1(2).Tag

    ElseIf Option1(3).Value Then
        corder = " order by " & Option1(3).Tag
        bVKM = True
    ElseIf Option1(4).Value Then
        corder = " order by " & Option1(4).Tag
        bVKM = True
    End If
    
    If Option2(0).Value Then
         corder = corder & " " & Option2(0).Tag
    ElseIf Option2(1).Value Then
         corder = corder & " " & Option2(1).Tag
    End If
    
    If bVKM = True Then
        corder = corder & " , Bestand "
        If Option2(0).Value Then
             corder = corder & " " & Option2(0).Tag
        ElseIf Option2(1).Value Then
             corder = corder & " " & Option2(1).Tag
        End If
    Else
        corder = corder & " , VKMENGE20 "
        If Option2(0).Value Then
             corder = corder & " " & Option2(0).Tag
        ElseIf Option2(1).Value Then
             corder = corder & " " & Option2(1).Tag
        End If
    End If
    
    cSQL = "insert into VKFAV select *  from FAVT "
    cSQL = cSQL & corder
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "FAVT", gdBase
    
    If Datendrin("VKFAV", gdBase) = False Then
        anzeige "rot", "Es wurden keine Daten ermittelt.", lblAnzeige
    Else
        anzeige "normal", "Druckvorschau wird erstellt...", lblAnzeige
        reportbildschirm "dWKL44", "aWKL44b"
        anzeige "normal", "Fertig", lblAnzeige
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseUmsatzDaten"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    If NewTableSuchenDBKombi("E44", gdApp) Then
        voreinstellungladen44
    End If
    
    
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladen44()
    On Error GoTo LOKAL_ERROR

    Dim rs As Recordset

    Set rs = gdApp.OpenRecordset("E44")
    If Not rs.EOF Then
        Option1(0).Value = rs!bo0
        Option1(1).Value = rs!bo1
        Option1(2).Value = rs!bo2
        Option1(3).Value = rs!bo3
        Option1(4).Value = rs!bo4
        Option2(0).Value = rs!bo5
        Option2(1).Value = rs!bo6
        
    End If
    rs.Close: Set rs = Nothing


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen44"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern44()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    Dim bo0 As Integer
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
    Dim bo5 As Integer
    Dim bo6 As Integer
   
    loeschNEW "E44", gdApp
    CreateTable "E44", gdApp

    bo0 = Option1(0).Value
    bo1 = Option1(1).Value
    bo2 = Option1(2).Value
    bo3 = Option1(3).Value
    bo4 = Option1(4).Value
    bo5 = Option2(0).Value
    bo6 = Option2(1).Value

    sSQL = "Insert into E44 ( bo0,bo1,bo2,bo3,bo4,bo5,bo6) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & "," & bo2 & "," & bo3 & "," & bo4
    sSQL = sSQL & " ," & bo5 & "," & bo6
    sSQL = sSQL & " )"
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern44"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

Dim ctmp As String
Dim lcount As Long
Dim sAuswahlfeld As String

If KeyCode = vbKeyF2 Then
    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = False
    
    Select Case Index
        Case Is = 2
            gF2Prompt.bMultiple = False
            gF2Prompt.cFeld = "LINR"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
            End If
            If gF2Prompt.cWahl <> "" Then
                Text1(Index).Text = gF2Prompt.cWahl
            End If
       

        Case 5
            ctmp = Text1(7).Text
            ctmp = Trim$(ctmp)
            If ctmp = "" Then
                ctmp = Text1(2).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    anzeige "Rot", "Bitte einen Lieferanten oder eine Marke angeben!", lblAnzeige
                    Text1(7).SetFocus
                    Exit Sub
                Else
                    sAuswahlfeld = "LINR"
                End If
            Else
                sAuswahlfeld = "MARKE"
            End If
            
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "LPZ"
            gF2Prompt.cWert = ctmp
            gF2Prompt.cEsFeld = sAuswahlfeld
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                List3.Visible = False
                List3.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List3.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount) & Space(50) & Right(gF2Prompt.cArray(lcount), 6)
                        End If
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left$(gF2Prompt.cArray(lcount), 3)
                        End If
                    End If
                Next lcount
            End If

        Case Is = 7
            gF2Prompt.cFeld = "MARKE"
            
            ctmp = Text1(2).Text 'Linr eventuell
            gF2Prompt.cEsFeld = ctmp
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
            End If

        End Select
        Text1(Index).SetFocus
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim ctmp As String
Dim lcount As Long
Dim sAuswahlfeld As String
    
    Select Case Index
        Case Is = 1
'            If Text1(7).Text <> "" And Text1(2).Text <> "" Then
'                anzeige "Rot", "Bitte nur einen Lieferant ODER eine Marke angeben!", lblanzeige
'                Exit Sub
'            End If
            ctmp = Text1(7).Text
            ctmp = Trim$(ctmp)
            If ctmp = "" Then
                ctmp = Text1(2).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    anzeige "Rot", "Bitte einen Lieferanten oder eine Marke angeben!", lblAnzeige
                    Text1(7).SetFocus
                    Exit Sub
                Else
                    sAuswahlfeld = "LINR"
                End If
            Else
                sAuswahlfeld = "MARKE"
            End If
            
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "LPZ"
            gF2Prompt.cWert = ctmp
            gF2Prompt.cEsFeld = sAuswahlfeld
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                List3.Clear
                For lcount = 0 To 100
                    If gF2Prompt.cArray(lcount) <> "" Then
                        List3.AddItem gF2Prompt.cArray(lcount)
                    End If
                Next lcount
            End If
       
        Case 0
            List3.Clear
        Case Is = 2
            Text1_KeyUp 2, vbKeyF2, 0
        Case 6
            Text1_KeyUp 7, vbKeyF2, 0
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
