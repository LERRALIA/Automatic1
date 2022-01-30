VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKLaf 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Datenbankbefehl"
   ClientHeight    =   8625
   ClientLeft      =   285
   ClientTop       =   1845
   ClientWidth     =   11910
   Icon            =   "frmWKLaf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      Height          =   285
      Index           =   5
      Left            =   7680
      TabIndex        =   31
      ToolTipText     =   "Leert wichtige Journaltabellen"
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
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
      Caption         =   "leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   285
      Index           =   4
      Left            =   8760
      TabIndex        =   30
      ToolTipText     =   "Löscht komplette unnötige Tabellen"
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
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
      Caption         =   "Ballast löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   285
      Index           =   3
      Left            =   10680
      TabIndex        =   29
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
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
      Caption         =   "Datum?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   6360
      TabIndex        =   28
      Top             =   7920
      Width           =   2655
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
      Caption         =   "Tabelle löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   6360
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   6480
      Width           =   5415
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   5775
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
      Height          =   1740
      Left            =   6360
      TabIndex        =   16
      Top             =   1680
      Width           =   5415
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
      Height          =   2370
      Left            =   6360
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   3720
      Width           =   5415
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   12
      Top             =   7920
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
      Left            =   120
      TabIndex        =   11
      Top             =   3960
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
      Caption         =   "Ausführen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Top             =   960
      Width           =   5415
      Begin VB.OptionButton Option1 
         Caption         =   "Kissapp.mdb"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Kissdata.mdb"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   285
      Index           =   6
      Left            =   7680
      TabIndex        =   32
      ToolTipText     =   "Leert wichtige Journaltabellen"
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
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
      Caption         =   "aufräumen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   285
      Index           =   7
      Left            =   8760
      TabIndex        =   33
      ToolTipText     =   "Löscht komplette unnötige Tabellen"
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
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
      Caption         =   "Werkseinstellung"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Tabelle enthält Indizes:"
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
      Left            =   6360
      TabIndex        =   27
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Anzahl"
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
      Left            =   9360
      TabIndex        =   26
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "letzte Datenbankbefehle:"
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
      TabIndex        =   24
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Datenbankbefehl"
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
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Anzahl"
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
      Left            =   9360
      TabIndex        =   18
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Anzahl"
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
      Left            =   9360
      TabIndex        =   17
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Tabelle enthält Felder:"
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
      Left            =   6360
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "vorhandene Tabellen:"
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
      Left            =   6360
      TabIndex        =   14
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "(Setze das Feld NAME auf MEIER, wo die Kundennummer 123456 lautet) - Achtung: Textfelder in einfache Anführungszeichen!"
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   8040
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "UPDATE Kunden SET NAME = 'MEIER' WHERE KUNDNR = 123456"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   7680
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "(Setze den Bestand aller Artikel auf 0)"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   7320
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "UPDATE Artikel SET BESTAND = 0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "(Setze den Kassen-VK auf den Preis des um 20 % reduzierten Listen-VK, wenn der Kassen-VK identisch zum Listen-VK ist)"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "UPDATE Artikel SET KVKPR1 = VKPR * 0.8 WHERE KVKPR1 = VKPR"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "Beispiele:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "UPDATE <Tabellenname> SET <Zielfeld> = <Neuer Wert> WHERE <Suchfeld> = <Bedingung>"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Aufbau des Datenbankbefehls für eine Änderung:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Datenbankbefehl eingeben:"
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
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "frmWKLaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbb As Database
Private Sub SpeicherSQL(cbet As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    cbet = SwapStr(cbet, Chr(39), "$")
    cbet = SwapStr(cbet, Chr(10), " ")
    cbet = SwapStr(cbet, Chr(13), " ")
    
    sSQL = "Insert into BEFEHLE (BTEXT,ZULETZT) Values "
    sSQL = sSQL & " ( '" & cbet & "  ' , Datevalue(now) )"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherSQL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
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
Private Sub BefehlAusfuehrenWKLaf(cSQL As String)
    On Error GoTo LOKAL_ERROR
    
   
    dbb.Execute cSQL, dbFailOnError
    MsgBox dbb.RecordsAffected & " Datensätze", vbInformation, "Winkiss Hinweis:"
    
    SpeicherSQL cSQL
    fuellelist3
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Ausführen eines SQL - Befehls"
    Fehler.gsFehlertext = cSQL
    
    Fehlermeldung1
    Exit Sub
    
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0     'Ausführen
            ctmp = Text1.Text
            ctmp = Trim$(ctmp)
            If ctmp = "" Then
                MsgBox "Bitte einen Datenbankbefehl eingeben!", vbInformation, "Winkiss Hinweis:"
                Text1.SetFocus
            Else
                BefehlAusfuehrenWKLaf ctmp
            End If
            
        Case Is = 1     'Schließen
        
            Dim sTabc As String
            sTabc = kassetabcheck(gdBase, Label2(0), Label2(1))
            
        
            If sTabc = "" Then
        
            Else
                MsgBox "Die Tabelle " & sTabc & " wurde nicht gefunden.", vbInformation, "Winkiss Hinweis:"
'                End
            End If
            Unload frmWKLaf
        Case 2
            tabelleDel
        Case Is = 3    'datum?
            Screen.MousePointer = 0
            frmWKLah.Show 1
            
        Case 4
        
            dlgPW.Show 1
                    
            If dlgPW.Back = True Then
                db_Ballast_Tabellen_del
                FuellList2 dbb
            Else
                MsgBox "Falsch"
            End If
        
           
            
            
        Case 5
        
            dlgPW.Show 1
                    
            If dlgPW.Back = True Then
                Journal_leeren
                FuellList2 dbb
            Else
                MsgBox "Falsch"
            End If
            
            
        Case 6
        
            dlgPW.Show 1
                    
            If dlgPW.Back = True Then
                aufräumen
                FuellList2 dbb
            Else
                MsgBox "Falsch"
            End If
            
        Case 7
        
            dlgPW.Show 1
                    
            If dlgPW.Back = True Then
                Werkseinstellungen
            Else
                MsgBox "Falsch"
            End If
            
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub tabelleDel()
On Error GoTo LOKAL_ERROR

    Dim cTabelle    As String
    Dim iRet        As Integer
    
    cTabelle = List2.list(List2.ListIndex)
    cTabelle = Left(cTabelle, 18)
    cTabelle = Trim(cTabelle)
    
    iRet = MsgBox("Soll die Tabelle '" & cTabelle & "' wirklich gelöscht werden?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
    If iRet = vbYes Then
        Text1.Text = "Drop Table " & cTabelle
        Command1_Click 0
        
'        FuellList2 dbb
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "tabelledel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List2_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cTabelle As String
    Dim tdTd As TableDef
    Dim lAnzFelder As Long
    Dim lcount As Long
    Dim lTyp As Long
    Dim lSize As Long
    Dim cFeldName As String
    Dim cFeldTyp As String
    Dim cFeldSize As String
    Dim cLBSatz As String
    
    Screen.MousePointer = 11
    
    List1.Clear
    List4.Clear
    
    cTabelle = List2.list(List2.ListIndex)
    cTabelle = Left(cTabelle, 18)
    cTabelle = Trim(cTabelle)
    
    Set tdTd = dbb.TableDefs(cTabelle)

    lAnzFelder = tdTd.Fields.Count
    For lcount = 0 To lAnzFelder - 1
        cFeldName = tdTd.Fields(lcount).name
        lTyp = tdTd.Fields(lcount).Type
        lSize = tdTd.Fields(lcount).Size
        cFeldSize = ""
        Select Case lTyp
            Case Is = dbDate
                cFeldTyp = "Datum"
            Case Is = dbText
                cFeldTyp = "Text"
                cFeldSize = Trim$(Str$(lSize))
            Case Is = dbMemo
                cFeldTyp = "Memofeld"
            Case Is = dbBoolean
                cFeldTyp = "Ja/Nein-Schalter"
            Case Is = dbInteger
                cFeldTyp = "Ganzzahl"
            Case Is = dbLong
                cFeldTyp = "Ganzzahl"
            Case Is = dbCurrency
                cFeldTyp = "Währung"
            Case Is = dbSingle
                cFeldTyp = "Kommazahl"
            Case Is = dbDouble
                cFeldTyp = "Kommazahl"
            Case Is = dbByte
                cFeldTyp = "Byte"
            Case Is = dbLongBinary
                cFeldTyp = "OLE-Objekt"
        End Select
        
        If cFeldSize <> "" Then
            cFeldTyp = cFeldTyp & " (" & cFeldSize & " Stellen) "
        End If
        
        cLBSatz = cFeldName & " - " & cFeldTyp & " "
        
        List1.AddItem cLBSatz
        
    Next lcount
    
    Label1(4).Caption = lAnzFelder - 1
    Label1(4).Refresh
    
    cTabelle = List2.list(List2.ListIndex)
    cTabelle = Left(cTabelle, 18)
    cTabelle = Trim(cTabelle)
    
    Set tdTd = dbb.TableDefs(cTabelle)

    lAnzFelder = tdTd.Indexes.Count
    For lcount = 0 To lAnzFelder - 1
        cFeldName = tdTd.Indexes(lcount).name
        cLBSatz = cFeldName
        List4.AddItem cLBSatz
        
    Next lcount
    
    Label1(6).Caption = lAnzFelder - 1
    Label1(6).Refresh

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3110 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "List2_Click"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    
    
    
    
    
    
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    If Not NewTableSuchenDBKombi("BEFEHLE", gdBase) Then
        CreateTable "BEFEHLE", gdBase
    End If
    
    Text1.Text = ""
    
    Option1_Click 0
    
    fuellelist3
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellelist3()
On Error GoTo LOKAL_ERROR

    Dim rsRS As Dao.Recordset
    Dim sSQL As String
    Dim cFeld As String
    Dim cLBSatz As String
    
    sSQL = "select * from befehle order by bnr desc"
    
    Screen.MousePointer = 11
    
    List3.Clear
    Set rsRS = gdBase.OpenRecordset(sSQL)
    If Not rsRS.EOF Then
        rsRS.MoveFirst
        Do While Not rsRS.EOF
            
            If Not IsNull(rsRS!zuletzt) Then
                cFeld = Format$(rsRS!zuletzt, "DD.MM.YY")
            Else
                cFeld = Space(8)
            End If
        
            cLBSatz = cFeld & Space$(2)
            
            If Not IsNull(rsRS!BTEXT) Then
                If Len(rsRS!BTEXT) >= 100 Then
                    cFeld = Left(rsRS!BTEXT, 95) & " ... "
                Else
                    cFeld = rsRS!BTEXT & Space(100 - Len(rsRS!BTEXT))
                End If
            Else
                cFeld = Space(100)
                
            End If
            cFeld = SwapStr(cFeld, "$", Chr(39))
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsRS!BNR) Then
                cFeld = rsRS!BNR
                cLBSatz = cLBSatz & Space(4) & cFeld
            End If
            
            List3.AddItem cLBSatz
            rsRS.MoveNext
        Loop
    End If
    rsRS.Close: Set rsRS = Nothing
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellelist3"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List3_Click()
On Error GoTo LOKAL_ERROR

    Dim rsRS As Recordset
    Dim sSQL As String
    Dim cFeld As String
    Dim lBnr As Long
    
    If List3.ListCount = 0 Then
        Exit Sub
    End If
    lBnr = CLng(Right(List3.list(List3.ListIndex), 4))
    
    sSQL = "select * from befehle where bnr = " & lBnr
    
    Screen.MousePointer = 11
    
    Text1.Text = ""
    
    Set rsRS = gdBase.OpenRecordset(sSQL)
    If Not rsRS.EOF Then
        If Not IsNull(rsRS!BTEXT) Then
            cFeld = rsRS!BTEXT
        Else
            cFeld = ""
        End If
        cFeld = SwapStr(cFeld, "$", Chr(39))
        Text1.Text = cFeld
    End If
    rsRS.Close: Set rsRS = Nothing
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
    
        Case 0
             Set dbb = gdBase
        
        Case 1
             Set dbb = gdApp
        
    End Select
    
    FuellList2 dbb
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellList2(db As Database)
On Error GoTo LOKAL_ERROR

    Dim lAnzTable       As Long
    Dim lcount          As Long
    Dim sTabname        As String
    Dim lMax            As Long
    Dim sSQL            As String
    Dim dateLastUpdate  As Date
    Dim rsRS            As Dao.Recordset
    
    
    loeschNEW "TABANZEIGE" & srechnertab, gdBase
    
    sSQL = "Create Table TABANZEIGE" & srechnertab & " "
    sSQL = sSQL & "(Tabname Text(100)"
    sSQL = sSQL & ", ANZDS LONG"
    sSQL = sSQL & ", LASTDATE Datetime"
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError
            
    
    db.TableDefs.Refresh
    lAnzTable = db.TableDefs.Count
    
    For lcount = 0 To lAnzTable - 1
        sTabname = db.TableDefs(lcount).name
        lMax = db.TableDefs(lcount).RecordCount
        dateLastUpdate = db.TableDefs(lcount).LastUpdated
        
        sSQL = "Insert into TABANZEIGE" & srechnertab & " "
        sSQL = sSQL & "(Tabname "
        sSQL = sSQL & ", ANZDS "
        sSQL = sSQL & ", LASTDATE "
        sSQL = sSQL & ") values ( '" & sTabname & "'," & lMax & ",'" & dateLastUpdate & "')"
        gdBase.Execute sSQL, dbFailOnError
        
'        List2.AddItem UCase(sTabname) & Space(30 - Len(sTabname)) & lMax & Space(12 - Len(CStr(lMax))) & " Datensätze"
    Next lcount
    
    
    Dim cLBSatz As String
    Dim cFeld As String
    
    
    List2.Clear
    
    sSQL = "select * from TABANZEIGE" & srechnertab & " "
    sSQL = sSQL & " order by  ANZDS desc "
    
    
    Set rsRS = gdBase.OpenRecordset(sSQL)
    If Not rsRS.EOF Then
        rsRS.MoveFirst
        Do While Not rsRS.EOF
            
            cLBSatz = ""
            
            If Not IsNull(rsRS!tabname) Then
                If Len(rsRS!tabname) >= 30 Then
                    cFeld = Left(rsRS!tabname, 27) & "..."
                Else
                    cFeld = rsRS!tabname & Space(30 - Len(rsRS!tabname))
                End If
            Else
                cFeld = Space(30)
            End If
            
            cLBSatz = cFeld
            
            
            
            
            If Not IsNull(rsRS!ANZDS) Then
                cFeld = Space(10 - Len(rsRS!ANZDS)) & rsRS!ANZDS
            Else
                cFeld = Space(10)
            End If
            
            cLBSatz = cLBSatz & Space(2) & cFeld
            
            
            
            
            
            If Not IsNull(rsRS!LASTDATE) Then
                cFeld = Format$(rsRS!LASTDATE, "DD.MM.YY")
            Else
                cFeld = Space(8)
            End If
        
            cLBSatz = cLBSatz & Space$(2) & cFeld
            
            List2.AddItem cLBSatz
            rsRS.MoveNext
        Loop
    End If
    rsRS.Close: Set rsRS = Nothing
    
    
    
'    db.TableDefs.Refresh
'    lAnzTable = db.TableDefs.Count
'
'    For lcount = 0 To lAnzTable - 1
'        sTabname = db.TableDefs(lcount).name
'        lMax = db.TableDefs(lcount).RecordCount
'
'        List2.AddItem UCase(sTabname) & Space(30 - Len(sTabname)) & lMax & Space(12 - Len(CStr(lMax))) & " Datensätze"
'    Next lcount
    Label1(3).Caption = lAnzTable - 1
    Label1(3).Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellList2"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub




