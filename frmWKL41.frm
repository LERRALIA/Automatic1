VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL41 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Artikelliste nach Artikelgruppen"
   ClientHeight    =   8625
   ClientLeft      =   390
   ClientTop       =   1740
   ClientWidth     =   11910
   Icon            =   "frmWKL41.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.Frame Frame6 
      Caption         =   "Ergebnisliste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   28
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6900
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   11415
      End
      Begin VB.ListBox List3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   11415
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   7680
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   1
         Left            =   9240
         TabIndex        =   18
         Top             =   7680
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   31
         Top             =   7920
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "von"
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   30
         Top             =   7920
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   7920
         TabIndex        =   29
         Top             =   7920
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Selektionsvorgabe"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   11655
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Height          =   3135
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   5175
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Rechts ausgerichtet
            BackColor       =   &H00C0C000&
            Caption         =   "Bestand > 0"
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
            Left            =   120
            TabIndex        =   53
            Top             =   2280
            Width           =   2820
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   0
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   1
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "JA"
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
            Left            =   2760
            TabIndex        =   2
            Top             =   1320
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "NEIN"
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
            Left            =   3840
            TabIndex        =   3
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Liste mit F2"
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
            Left            =   3960
            TabIndex        =   52
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Liste mit F2"
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
            Left            =   3960
            TabIndex        =   51
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "von Artikelgruppen-Nummer..."
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
            TabIndex        =   27
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "bis Artikelgruppen-Nummer..."
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
            TabIndex        =   26
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "EK-Preise drucken..."
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
            TabIndex        =   25
            Top             =   1320
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C000&
         Caption         =   "Sortierung"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   5400
         TabIndex        =   23
         Top             =   240
         Width           =   6135
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Artikelgruppennummer, Alphabet"
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
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   4335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Artikelgruppennummer, Lieferantennummer"
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
            TabIndex        =   5
            Top             =   720
            Width           =   4335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Artikelgruppennummer, Artikelnummer"
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
            TabIndex        =   6
            Top             =   1080
            Width           =   4335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Verkaufsdatum"
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
            TabIndex        =   7
            Top             =   1440
            Width           =   4335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Lieferantennummer, Linie, Lieferantenbestellnummer"
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
            TabIndex        =   8
            Top             =   1800
            Visible         =   0   'False
            Width           =   5895
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Lieferantennummer, Artikelnummer"
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
            TabIndex        =   9
            Top             =   2160
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Verkaufsdatum"
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
            TabIndex        =   10
            Top             =   2520
            Visible         =   0   'False
            Width           =   4335
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C000&
         Caption         =   "Liniennummern"
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
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   0
            Left            =   2760
            MaxLength       =   3
            TabIndex        =   11
            Text            =   "Text2"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   1
            Left            =   2760
            MaxLength       =   3
            TabIndex        =   12
            Text            =   "Text2"
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "von Linie..."
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
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "bis Linie..."
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
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   2055
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   9240
         TabIndex        =   14
         Top             =   4320
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   5400
         TabIndex        =   13
         Top             =   4320
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   2295
      Left            =   0
      TabIndex        =   32
      Top             =   6000
      Width           =   12015
      Begin sevCommand3.Command Command0 
         Height          =   975
         Index           =   15
         Left            =   10680
         TabIndex        =   50
         Top             =   1200
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   1
         Left            =   1080
         TabIndex        =   46
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   2
         Left            =   2040
         TabIndex        =   45
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   3
         Left            =   3000
         TabIndex        =   44
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   4
         Left            =   3960
         TabIndex        =   43
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   5
         Left            =   4920
         TabIndex        =   42
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   6
         Left            =   5880
         TabIndex        =   41
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   7
         Left            =   6840
         TabIndex        =   40
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   8
         Left            =   7800
         TabIndex        =   39
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   9
         Left            =   8760
         TabIndex        =   38
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   10
         Left            =   9720
         TabIndex        =   37
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   11
         Left            =   10680
         TabIndex        =   36
         Top             =   240
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   12
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
         Height          =   975
         Index           =   13
         Left            =   3960
         TabIndex        =   34
         Top             =   1200
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "<<<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   975
         Index           =   14
         Left            =   5880
         TabIndex        =   33
         Top             =   1200
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   ">>>"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   49
         Top             =   1440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   48
         Top             =   1800
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikelliste nach Artikelgruppen"
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
      Index           =   0
      Left            =   120
      TabIndex        =   54
      Top             =   120
      Width           =   10815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "ARTHEAD", gdBase
    loeschNEW "ARTDRUCK", gdBase
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
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0 To 9         'Ziffern
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    Text1(Val(Label0(1).Caption)).Text = Text1(Val(Label0(1).Caption)).Text & Command0(Index).Caption
                    Text1(Val(Label0(1).Caption)).SetFocus
                ElseIf Label0(0).Caption = "Text2" Then
                    Text2(Val(Label0(1).Caption)).Text = Text2(Val(Label0(1).Caption)).Text & Command0(Index).Caption
                    Text2(Val(Label0(1).Caption)).SetFocus
                End If
            End If
        Case Is = 10        'Backspace
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    If Len(Text1(Val(Label0(1).Caption)).Text) > 0 Then
                        Text1(Val(Label0(1).Caption)).Text = Left(Text1(Val(Label0(1).Caption)).Text, Len(Text1(Val(Label0(1).Caption)).Text) - 1)
                    End If
                    Text1(Val(Label0(1).Caption)).SetFocus
                ElseIf Label0(0).Caption = "Text2" Then
                    If Len(Text2(Val(Label0(1).Caption)).Text) > 0 Then
                        Text2(Val(Label0(1).Caption)).Text = Left(Text2(Val(Label0(1).Caption)).Text, Len(Text2(Val(Label0(1).Caption)).Text) - 1)
                    End If
                    Text2(Val(Label0(1).Caption)).SetFocus
                End If
            End If
        Case Is = 11        'Clear
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    Text1(Val(Label0(1).Caption)).Text = ""
                    Text1(Val(Label0(1).Caption)).SetFocus
                ElseIf Label0(0).Caption = "Text2" Then
                    Text2(Val(Label0(1).Caption)).Text = ""
                    Text2(Val(Label0(1).Caption)).SetFocus
                End If
            End If
        Case Is = 12        'F2
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    Text1_KeyUp Val(Label0(1).Caption), vbKeyF2, 0
                ElseIf Label0(0).Caption = "Text2" Then
                    'Text2_KeyUp Val(Label0(1).Caption), vbKeyF2, 0
                End If
            End If
        Case Is = 13        'vorheriges Element
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    Label0(1).Caption = Val(Label0(1).Caption) - 1
                    If Val(Label0(1).Caption) < 0 Then
                        Label0(1).Caption = "0"
                    End If
                    Text1(Val(Label0(1).Caption)).SetFocus
                End If
            End If
                        
        Case Is = 14        'nachfolgendes Element
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    Label0(1).Caption = Val(Label0(1).Caption) + 1
                    If Val(Label0(1).Caption) > 2 Then
                        Label0(1).Caption = "2"
                    End If
                    Text1(Val(Label0(1).Caption)).SetFocus
                End If
            End If
        Case Is = 15        'Punkt für Datum
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    If Val(Label0(1).Caption) = 2 Then
                        Text1(Val(Label0(1).Caption)).Text = Text1(Val(Label0(1).Caption)).Text & Command0(Index).Caption
                        Text1(Val(Label0(1).Caption)).SetFocus
                    Else
                        Text1(Val(Label0(1).Caption)).SetFocus
                    End If
                ElseIf Label0(0).Caption = "Text2" Then
                    Text2(Val(Label0(1).Caption)).SetFocus
                End If
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
        
    Screen.MousePointer = 11
        
    Select Case Index
        Case Is = 0
            iRet = fnPruefeDialogWKL41%()
            Select Case iRet
                Case Is = 0
                    LeseDatenWKL41

                    Frame6.Visible = True
                    Frame0.Visible = False
                    Frame1.Enabled = False
                Case Is = 1     'keine AGN
                    MsgBox "Bitte eine Artikelgruppennummer angeben!", vbInformation, "Winkiss Hinweis:"
                    Text1(0).SetFocus
                    Screen.MousePointer = 0
                    Exit Sub
                
            End Select
        Case Is = 1
            Unload frmWKL41
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LeseDatenWKL41()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL       As String
    Dim rsrs       As Recordset
    Dim cLiNr1     As String
    Dim cLiNr2     As String
    
    Dim lcount     As Long
    Dim lOrderBy   As Long
    Dim ctmp       As String
    Dim cEkPr      As String
    Dim dWert      As Double
    Dim dWert2     As Double
    Dim dErgebnis  As Double
    Dim cFeld      As String
    Dim cLBSatz    As String
    
    'Summenfelder
    Dim dGesBestand        As Double
    Dim dGesVkLfdMo        As Double
    Dim dGesUmsLfdMo       As Double
    Dim dGesVkVorMo        As Double
    Dim dGesUmsVorMo       As Double
    Dim dErgebnis2         As Double
    ReDim cOrderBy(0 To 6) As String
    
    cOrderBy(0) = "order by AGN, BEZEICH "
    cOrderBy(1) = "order by AGN, LINR "
    cOrderBy(2) = "order by AGN, ARTNR "
    cOrderBy(3) = "order by VKDATUM "
    cOrderBy(4) = ""
    cOrderBy(5) = ""
    cOrderBy(6) = ""
    
    cLiNr1 = Text1(0).Text
    cLiNr2 = Text1(1).Text
    cLiNr1 = Trim$(cLiNr1)
    cLiNr2 = Trim$(cLiNr2)
    
    For lcount = 0 To 6
        If Option2(lcount).Value = True Then
            lOrderBy = lcount
            Exit For
        End If
    Next lcount
    
    cSQL = "Select * from ARTIKEL where AGN >= " & cLiNr1 & " and AGN <= " & cLiNr2 & " "
    If Check1.Value = 1 Then
        cSQL = cSQL & " and bestand > 0 "
    End If
    cSQL = cSQL & cOrderBy(lOrderBy)
    
'    MsgBox cSQL
    
    List3.Clear
    List4.Clear
    
    If Option1(0).Value = True Then
        cEkPr = " EK-Preis"
    Else
        cEkPr = ""
    End If
    
    ctmp = "Art.Nr. Artikel-Bezeichnung                 Linie   AGN EAN           E R LiefNr LiefBestNr    "
    
    If Option1(0).Value = True Then
        ctmp = ctmp & cEkPr & " "
    End If
    
    ctmp = ctmp & " VK-Preis Bestand VK-Menge LetzterVK Mon.Umsatz"
    
    List3.AddItem ctmp
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    dErgebnis2 = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            
            cFeld = SwapStr(cFeld, "'", "")
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " " & Space$(3)
            
            If Not IsNull(rsrs!LPZ) Then
                cFeld = rsrs!LPZ
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!AGN) Then
                cFeld = rsrs!AGN
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!EAN) Then
                cFeld = rsrs!EAN
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(13 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!ETIMERK) Then
                cFeld = rsrs!ETIMERK
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(1 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!RKZ) Then
                cFeld = rsrs!RKZ
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(1 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!linr) Then
                cFeld = rsrs!linr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!LIBESNR) Then
                cFeld = rsrs!LIBESNR
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(13 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Option1(0).Value = True Then
                If Not IsNull(rsrs!lekpr) Then
                    cFeld = rsrs!lekpr
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cFeld = fnMoveComma2Point$(cFeld)
                dWert = Val(cFeld)
                dWert2 = dWert
                cFeld = Format$(dWert, "#####0.00")
                cFeld = Space$(9 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
            End If
            
            If Not IsNull(rsrs!KVKPR1) Then
                cFeld = rsrs!KVKPR1
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = fnMoveComma2Point$(cFeld)
            dWert = Val(cFeld)
            cFeld = Format$(dWert, "#####0.00")
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!BESTAND) Then
                cFeld = rsrs!BESTAND
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = fnMoveComma2Point$(cFeld)
            dWert = Val(cFeld)
            
            '//LEKPR * BESTAND
            dErgebnis = dWert2 * dWert
            dErgebnis2 = dErgebnis2 + dErgebnis
            dGesBestand = dGesBestand + dWert
            cFeld = Format$(dWert, "######0")
            cFeld = Space$(7 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!VKMENGE) Then
                cFeld = rsrs!VKMENGE
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = fnMoveComma2Point$(cFeld)
            dWert = Val(cFeld)
            dGesVkLfdMo = dGesVkLfdMo + dWert
            cFeld = Format$(dWert, "######0")
            cFeld = Space$(7 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!VKDATUM) Then
                cFeld = rsrs!VKDATUM
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!MOPREIS) Then
                cFeld = rsrs!MOPREIS
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = fnMoveComma2Point$(cFeld)
            dWert = Val(cFeld)
            dGesUmsLfdMo = dGesUmsLfdMo + dWert
            cFeld = Format$(dWert, "#####0.00")
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            List4.AddItem cLBSatz
            rsrs.MoveNext
        Loop
        
        List4.AddItem String$(160, "-")
        
        If Option1(1).Value = True Then
            cLBSatz = Space$(80) & "Gesamtbestand ="
            
            cLBSatz = cLBSatz & Space$(8)
            cFeld = Format$(dGesBestand, "########0")
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            cFeld = Format$(dGesVkLfdMo, "#####0")
            cFeld = Space$(7 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " " & Space$(11)
            
            cFeld = Format$(dGesUmsLfdMo, "#####0.00")
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
        Else
            cLBSatz = Space$(86) & "EK-WERT ="
            cFeld = Format$(dErgebnis2, "#####0.00")
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            cLBSatz = cLBSatz & Space$(8)
            cFeld = Format$(dGesBestand, "########0")
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            cFeld = Format$(dGesVkLfdMo, "#####0")
            cFeld = Space$(7 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " " & Space$(11)
            
            cFeld = Format$(dGesUmsLfdMo, "#####0.00")
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
        End If
        
        List4.AddItem cLBSatz
        Label3(2).Caption = List4.ListCount
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseDatenWKL41"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub LeereDialogWKL41()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text2(0).Text = ""
    Text2(1).Text = ""
    Option1(0).Value = True
    Option2(0).Value = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL41"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeDialogWKL41%()
    On Error GoTo LOKAL_ERROR
    
    Dim cLiNr1 As String
    Dim cLiNr2 As String
    Dim ctmp As String
    Dim dWert As Double
    
    cLiNr1 = Text1(0).Text
    cLiNr2 = Text1(1).Text
    cLiNr1 = Trim$(cLiNr1)
    cLiNr2 = Trim$(cLiNr2)
    
    If cLiNr1 = "" And cLiNr2 = "" Then
        fnPruefeDialogWKL41% = 1
        Exit Function
    End If
    
    If cLiNr1 = "" And cLiNr2 <> "" Then
        Text1(0).Text = cLiNr2
    End If
    
    If cLiNr1 <> "" And cLiNr2 = "" Then
        Text1(1).Text = cLiNr1
    End If
    
    cLiNr1 = Text1(0).Text
    cLiNr2 = Text1(1).Text
    cLiNr1 = Trim$(cLiNr1)
    cLiNr2 = Trim$(cLiNr2)
    
    If Val(cLiNr1) > Val(cLiNr2) Then
        Text1(0).Text = cLiNr2
        Text1(1).Text = cLiNr1
    End If
    
    fnPruefeDialogWKL41% = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDialogWKL41%"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub DruckeErgebnisListeWKL41()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim lcount As Long
    Dim lAnzSatz As Long
    Dim cLBSatz As String
    
    loeschNEW "ARTHEAD", gdBase
    loeschNEW "ARTDRUCK", gdBase
    
    cSQL = "Create Table ARTHEAD (SCHLUESSEL DOUBLE, LISTTEXT TEXT(180))"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create INDEX SCHLUESSEL on ARTHEAD (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Table ARTDRUCK (SCHLUESSEL DOUBLE, LISTTEXT TEXT(180))"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create INDEX SCHLUESSEL on ARTDRUCK (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cLBSatz = List3.list(0)
    cSQL = "Insert into ARTHEAD (SCHLUESSEL, LISTTEXT) values (1, '" & cLBSatz & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    lAnzSatz = List4.ListCount
    
    Label3(2).Caption = Format$(lAnzSatz, "###,##0")
    Label3(2).Refresh
    For lcount = 0 To lAnzSatz - 1
        Label3(0).Caption = Format$(lcount + 1, "###,##0")
        Label3(0).Refresh
        cLBSatz = List4.list(lcount)
        
        cSQL = "Insert into ARTDRUCK (SCHLUESSEL, LISTTEXT) values (1, '" & cLBSatz & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next lcount
        
    reportbildschirm "WKL001", "aWKL41"

    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeErgebnisListeWKL41"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub



Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0     'Ergebnisliste drucken
            DruckeErgebnisListeWKL41
        Case Is = 1     'Ergebnisliste schließen
            Frame0.Visible = True
            Frame1.Enabled = True
            Frame6.Visible = False
    End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnz As Long
    Dim lcount As Long
    
    Screen.MousePointer = 11
    
    PositionierenWKL41
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift(0)
    
    LeereDialogWKL41
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub PositionierenWKL41()
    On Error GoTo LOKAL_ERROR
    
    With Frame6
        .Height = 8295
        .Top = 0
        .Width = 11655 '11895
        .Left = 120
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL41"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    Label0(0).Caption = Text1(Index).name
    Label0(1).Caption = Trim$(Str$(Index))
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Or KeyCode = vbKeyF4 Then
        
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 0
                gF2Prompt.cFeld = "AGN"
                
            Case Is = 1
                gF2Prompt.cFeld = "AGN"
                    
        End Select
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
        End If
        
        If gF2Prompt.cWahl <> "" Then
            Text1(Index).Text = gF2Prompt.cWahl
        End If
        Text1(Index).SetFocus

    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Artikelgruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub


