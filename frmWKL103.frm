VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL103 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Favoritenliste"
   ClientHeight    =   8625
   ClientLeft      =   3105
   ClientTop       =   2055
   ClientWidth     =   11910
   Icon            =   "frmWKL103.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
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
      Height          =   3615
      Left            =   0
      TabIndex        =   16
      Top             =   720
      Width           =   12015
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefbestNr"
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
         TabIndex        =   37
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Index           =   5
         Left            =   7440
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelbezeichnung"
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
         Left            =   7440
         TabIndex        =   34
         Top             =   600
         Width           =   2055
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
         Index           =   2
         Left            =   7440
         TabIndex        =   33
         Top             =   960
         Width           =   1455
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
         Height          =   375
         Index           =   2
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2880
         Width           =   1695
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   15
         Top             =   720
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
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   14
         Top             =   120
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
         Caption         =   "Suche Daten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   3375
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaxLength       =   3
            TabIndex        =   5
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00008080&
            BackStyle       =   0  'Transparent
            Caption         =   "Produktlinien (Liste mit F2, wenn LiNr vorh.)"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label Label1 
            BackColor       =   &H00008080&
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00008080&
            BackStyle       =   0  'Transparent
            Caption         =   "( Auswahlliste mit Taste F2 )"
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   25
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackColor       =   &H00008080&
            BackStyle       =   0  'Transparent
            Caption         =   "Lieferanten-Nummer (leer = alle):"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C000&
         Caption         =   "Suchbedingung 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7800
         TabIndex        =   21
         Top             =   1440
         Width           =   3975
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            Text            =   "0"
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "mindestens"
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "maximal"
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   12
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "exakt"
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   13
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00008080&
            BackStyle       =   0  'Transparent
            Caption         =   "vorhandene Stückzahl im Geschäft:"
            Height          =   735
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C000&
         Caption         =   "Suchbedingung 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3600
         TabIndex        =   19
         Top             =   1440
         Width           =   4095
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "exakt"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   9
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "maximal"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "mindestens"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   6
            Text            =   "0"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00008080&
            BackStyle       =   0  'Transparent
            Caption         =   "verkaufte Stückzahl im Zeitraum:"
            Height          =   735
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Stückzahlen"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz"
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
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
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
         Height          =   375
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
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
         Height          =   375
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1695
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   20
         Left            =   3000
         TabIndex        =   38
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   21
         Left            =   3000
         TabIndex        =   39
         ToolTipText     =   "Kalender"
         Top             =   720
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
         PictureAlign    =   2
         Version3        =   -1  'True
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
         TabIndex        =   36
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "eingekauft worden sind, werden nicht berücksichtigt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   7440
         TabIndex        =   30
         Top             =   2880
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   " Artikel, die nach dem "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   29
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Datum von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Datum bis:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1095
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
      TabIndex        =   32
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
      TabIndex        =   31
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
Attribute VB_Name = "frmWKL103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPrueD          As Integer
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        
        Case Is = 20        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            Text1(1).SetFocus
            
        Case Is = 21        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            'fertig
        End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "vkpro1", gdBase
    loeschNEW "vkproko", gdBase
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
    
    Dim iRet As Integer
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0
            iRet = fnPruefeEingabeWKL44%()
            Select Case iRet
                Case Is = 0
                    LeseUmsatzDatenWKL44
                Case Is = 1     'VON-Feld nicht gefüllt
                    MsgBox "Das Feld 'Datum von' muß gefüllt sein!", vbCritical, "STOP!"
                    Text1(0).SetFocus
                Case Is = 2     'BIS-Feld nicht gefüllt
                    MsgBox "Das Feld 'Datum bis' muß gefüllt sein!", vbCritical, "STOP!"
                    Text1(1).SetFocus
                
                Case Is = 11    'VON-Feld ist kein Datum
                    MsgBox "Das eingegebene Datum im Feld 'Datum von' ist ungültig!", vbCritical, "STOP!"
                    Text1(0).SetFocus
                
                Case Is = 21    'BIS-Feld ist kein Datum
                    MsgBox "Das eingegebene Datum im Feld 'Datum bis' ist ungültig!", vbCritical, "STOP!"
                    Text1(1).SetFocus
                
                Case Is = 99    'VON ist größer als BIS
                    MsgBox "Das eingegebene Datum im Feld 'Datum von' ist größer als das Datum im Feld 'Datum bis'!", vbCritical, "STOP!"
                    Text1(0).SetFocus
                
            End Select
            
        Case Is = 1
            iRet = MsgBox("Möchten Sie wirklich die Favoritenliste beenden?", vbQuestion + vbYesNo, "ENDE")
            If iRet = vbYes Then
                Unload frmWKL103
            End If
            
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
Private Sub LeseUmsatzDatenWKL44()
    On Error GoTo LOKAL_ERROR
    
    Dim cVon        As String
    Dim cBis        As String
    Dim cLastEK     As String
    
    Dim lVon        As Long
    Dim lBis        As Long
    Dim lLastEk     As Long
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim lAnzRecords As Long
    Dim lAktRecord As Long
    Dim ctmp As String
    Dim dWert As Double
    Dim cOrderBy As String
    
    Dim dEinkauf As Double
    Dim dErtrag As Double
    Dim dAufschlag As Double
    
    Dim dUmsatz As Double
    Dim dAnzKunden As Double
    
    Dim dVkPr As Double
    Dim dEkpr As Double
    Dim dAnz As Double
    
    Dim cLinr As String
    Dim clpz As String
    Dim cArtNr As String
    
    Dim dSummeUmsatz As Double
    Dim lSummeAnzahl As Long
    
    Dim lVKVorgabe  As Long
    Dim lBsVorgabe  As Long
    Dim cBsTyp      As String
    Dim cVkTyp      As String
    
    Dim cBestand As String
    
    lblanzeige.Caption = "Daten werden ermittelt..."
    lblanzeige.Refresh
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    
    If Text1(2).Text <> "" Then
        cLastEK = Text1(2).Text
        lLastEk = DateValue(cLastEK)
        cLastEK = Trim$(Str$(lLastEk))
    End If
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    cLinr = Text3.Text
    cLinr = Trim$(Str$(Val(cLinr)))
    If cLinr = "0" Then
        cLinr = "*"
    End If
    
    clpz = Text4.Text
    clpz = Trim$(Str$(Val(clpz)))
    If clpz = "0" Then
        clpz = "*"
    End If
    
    cBestand = Text2(1).Text
    cBestand = Trim$(cBestand)
        
    '***************************************
    '* vorhandene TEMP-Tabelle löschen
    '***************************************
    
    loeschNEW "vkpro1", gdBase

    '***************************************
    '* TEMP-Tabelle neu erzeugen
    '***************************************
    
    cSQL = "Create Table vkpro1 (ARTNR Double"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", UMSATZ Double"
    cSQL = cSQL & ", STUECK Double"
    cSQL = cSQL & ", BESTAND Double"
    cSQL = cSQL & ", EK_WERT Double"
    cSQL = cSQL & ", VK_WERT Double"
    cSQL = cSQL & ", LINR Long"
    cSQL = cSQL & ", LPZ Long"
    cSQL = cSQL & ")"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    '***************************************
    '* TEMP-Tabelle füllen
    '***************************************
    
    cSQL = "Insert into vkpro1 "
    cSQL = cSQL & "Select "
    cSQL = cSQL & "ARTIKEL.ARTNR "
    cSQL = cSQL & ", ARTIKEL.BEZEICH "
    cSQL = cSQL & ", ARTIKEL.LIBESNR "
    cSQL = cSQL & ", sum(kassjour.preis) as UMSATZ"
    cSQL = cSQL & ", sum(kassjour.menge) as STUECK"
    cSQL = cSQL & ", ARTIKEL.BESTAND "
    cSQL = cSQL & ", (ARTIKEL.BESTAND * ARTIKEL.EKPR) as EK_WERT "
    cSQL = cSQL & ", (ARTIKEL.BESTAND * ARTIKEL.KVKPR1) as VK_WERT "
    cSQL = cSQL & ", ARTIKEL.LINR "
    cSQL = cSQL & ", ARTIKEL.LPZ "
    cSQL = cSQL & " from ARTIKEL inner join Kassjour on "
    cSQL = cSQL & " ARTIKEL.artnr = Kassjour.artnr where "
    
    cSQL = cSQL & " Kassjour.ADATE Between " & Trim$(Str$(lVon)) & " And " & Trim$(Str$(lBis)) & " and "
    
    If cLinr = "*" Then
    Else
        cSQL = cSQL & " ARTIKEL.LINR = " & cLinr & " and "
    End If
    
    If clpz = "*" Then
    Else
        cSQL = cSQL & " ARTIKEL.LPZ = " & clpz & " and "
    End If
    
    
    
    If cBestand <> "" Then
        cSQL = cSQL & " ARTIKEL.BESTAND "
        If Option2(3).Value = True Then
            cSQL = cSQL & " >= "
        End If
        If Option2(4).Value = True Then
            cSQL = cSQL & " <= "
        End If
        If Option2(5).Value = True Then
            cSQL = cSQL & " = "
        End If
        
        cSQL = cSQL & cBestand
    End If
    cSQL = cSQL & " group by ARTIKEL.ARTNR, ARTIKEL.BEZEICH, ARTIKEL.KVKPR1 "
    cSQL = cSQL & ", ARTIKEL.LIBESNR,ARTIKEL.EKPR, ARTIKEL.BESTAND, ARTIKEL.LINR, ARTIKEL.LPZ "
    gdBase.Execute cSQL, dbFailOnError
    
    
    '*******************************************************************************
    ' Jetzt alles aus vkpro1 rauswerfen, was nicht den Suchkriterien entspricht! *
    '*******************************************************************************
                
    ctmp = Text2(0).Text
    ctmp = Trim$(ctmp)
                
    cSQL = "Delete from vkpro1 where STUECK "
    
    If Option2(0).Value = True Then
        cSQL = cSQL & "< "
    ElseIf Option2(1).Value = True Then
        cSQL = cSQL & "> "
    ElseIf Option2(2).Value = True Then
        cSQL = cSQL & "<> "
    End If
    cSQL = cSQL & ctmp & " "
    gdBase.Execute cSQL, dbFailOnError
    
    If Text1(2).Text <> "" Then
    
        cSQL = "Delete from vkpro1 where artnr in (Select artnr from  Zugang where zugang.adate > " & cLastEK & ")"
        gdBase.Execute cSQL, dbFailOnError
        
    End If

    loeschNEW "vkpro", gdBase
    cSQL = "Select * into Vkpro from vkpro1 "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "vkproko", gdBase
    cSQL = "Create Table vkproko "
    cSQL = cSQL & "( DAT_VON Text(10)"
    cSQL = cSQL & ", DAT_BIS Text(10)"
    cSQL = cSQL & ", LINR Text(6)"
    cSQL = cSQL & ", VK_VORGABE Long"
    cSQL = cSQL & ", VK_TYP Text(10)"
    cSQL = cSQL & ", BS_VORGABE Long"
    cSQL = cSQL & ", BS_TYP Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cVon = Text1(0).Text
    cBis = Text1(1).Text
    cLinr = Trim$(Text3.Text)
    If Trim$(Text2(0).Text) = "" Then
        lVKVorgabe = -999999
        cVkTyp = ""
    Else
        lVKVorgabe = Val(Text2(0).Text)
        If Option2(0).Value Then
            cVkTyp = "mindestens"
        ElseIf Option2(1).Value Then
            cVkTyp = "maximal"
        ElseIf Option2(2).Value Then
            cVkTyp = "exakt"
        End If
    End If

    If Trim$(Text2(1).Text) = "" Then
        lBsVorgabe = -999999
        cBsTyp = "mindestens"
    Else
        lBsVorgabe = Val(Text2(1).Text)
        If Option2(3).Value Then
            cBsTyp = "mindestens"
        ElseIf Option2(4).Value Then
            cBsTyp = "maximal"
        ElseIf Option2(5).Value Then
            cBsTyp = "exakt"
        End If
    End If

    cSQL = "Select * from vkproko"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    rsrs.AddNew
    rsrs!DAT_VON = cVon
    rsrs!DAT_BIS = cBis
    If cLinr <> "" Then
        rsrs!linr = cLinr
    Else
        rsrs!linr = "alle"
    End If
    rsrs!VK_VORGABE = lVKVorgabe
    rsrs!VK_TYP = cVkTyp
    rsrs!BS_VORGABE = lBsVorgabe
    rsrs!BS_TYP = cBsTyp
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    If SpalteInTabellegefundenNEW("vkpro", "EAN", gdBase) = False Then
        SpalteAnfuegenNEW "vkpro", "EAN", "Text(13)", gdBase
    End If
    
    cSQL = "Update vkpro inner join Artikel on vkpro.ARTNR = Artikel.Artnr "
    cSQL = cSQL & " set  vkpro.ean = ARTIKEL.ean "
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("vkpro", dbOpenTable)
    If rsrs.EOF Then
        lblanzeige.Caption = "Es wurden keine Daten ermittelt."
        lblanzeige.Refresh
    Else
        If Option1(0).Value = True Then                         'Umsatz
            Sortierung 0
            reportbildschirm "dWKL44", "aWKL44"
        ElseIf Option1(1).Value = True Then                     'Stückzahl
            Sortierung 1
            reportbildschirm "dWKL44", "aWKL44"
        ElseIf Option1(2).Value = True Then                     'Bestand
            Sortierung 2
            reportbildschirm "dWKL44", "aWKL44"
        ElseIf Option1(3).Value = True Then                     'Libesnr
            Sortierung 3
            reportbildschirm "dWKL44a", "aWKL44a"
        ElseIf Option1(4).Value = True Then                     'Bez
            Sortierung 4
            reportbildschirm "dWKL44a", "aWKL44a"
        ElseIf Option1(5).Value = True Then                     'artnr
            Sortierung 5
            reportbildschirm "dWKL44", "aWKL44"
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseUmsatzDatenWKL44"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeEingabeWKL44%()
    On Error GoTo LOKAL_ERROR
    
    Dim cVon As String
    Dim cBis As String
    Dim lVon As Long
    Dim lBis As Long
    Dim ctmp As String
    Dim lTmp As Long
    
    fnPruefeEingabeWKL44% = 0
    
    cVon = Text1(0).Text
    cVon = Trim$(cVon)
    If cVon = "" Then
        fnPruefeEingabeWKL44% = 1
        Exit Function
    End If
    
    cBis = Text1(1).Text
    cBis = Trim$(cBis)
    If cBis = "" Then
        fnPruefeEingabeWKL44% = 2
        Exit Function
    End If
    
    If Not IsDate(cVon) Then
        fnPruefeEingabeWKL44% = 11
        Exit Function
    End If
        
    If Not IsDate(cBis) Then
        fnPruefeEingabeWKL44% = 21
        Exit Function
    End If
        
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    If lVon > lBis Then
        fnPruefeEingabeWKL44% = 99
        Exit Function
    End If
        
    Text1(0).Text = Format$(lVon, "DD.MM.YY")
    Text1(1).Text = Format$(lBis, "DD.MM.YY")
        
    ctmp = Text2(0).Text
    ctmp = Trim$(ctmp)
    lTmp = Val(ctmp)
    ctmp = Trim$(Str$(lTmp))
    Text2(0).Text = ctmp
    
    ctmp = Text2(1).Text
    ctmp = Trim$(ctmp)
    lTmp = Val(ctmp)
    ctmp = Trim$(Str$(lTmp))
    Text2(1).Text = ctmp
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWKL44"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim cMM As String
    Dim cYY As String
    
    Screen.MousePointer = 11
    
    PositionierenWKL44
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Text1(0).Text = Format$("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YY")
    Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
    Text1(1).Text = DateValue(Now) - 1
    
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
Private Sub PositionierenWKL44()
    On Error GoTo LOKAL_ERROR
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL44"
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
Private Sub Text2_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text2(Index).BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cValid = "1234567890-" & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Beep
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Text3_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = "LINR"
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
        End If
        
        If gF2Prompt.cWahl <> "" Then
            Text3.Text = gF2Prompt.cWahl
        End If
        Text3.SetFocus
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.BackColor = vbWhite
    If Trim(Text3.Text) = "" Then
        Label1(6).Caption = ""
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        If Text3.Text = "" Then
            MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
            Exit Sub
        End If
        
        gF2Prompt.cFeld = "LPZ"
        gF2Prompt.cWert = Text3.Text
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
        End If
        
        If gF2Prompt.cWahl <> "" Then
            Text4.Text = gF2Prompt.cWahl
        End If
        Text4.SetFocus
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Favoritenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub





