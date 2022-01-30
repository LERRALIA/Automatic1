VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL183 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "MDE betanken"
   ClientHeight    =   8610
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   11910
   Icon            =   "frmWKL183.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check3 
      Caption         =   "nur mit Mindestbestand oder Bestand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   25
      Top             =   4680
      Width           =   4815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "nur mit Mindestbestand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "nur mit Bestand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   3375
   End
   Begin VB.CheckBox Check6 
      Caption         =   "nur geführte Artikel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Frame Frame7 
      BorderStyle     =   0  'Kein
      Height          =   5655
      Left            =   8040
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   3615
      Begin VB.OptionButton Option1 
         Caption         =   "3 Zeilen für Lieferanten, EK und VPE + Bestände, Kassenverkaufspreis, letzter Verkauf und die Verkaufsmenge der letzten 30 Tage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 Zeilen für Lieferanten, EK und VPE + Kassenverkaufspreis, letzter Zugang und die Zugangsmenge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 Zeilen für Lieferanten, EK und VPE + Bestände, Mindestbestände, letzter Verkauf und die Verkaufsmenge der letzten 30 Tage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 Zeilen für Lieferanten, EK und VPE + Bestände und Mindestbestände von Filiale 1 und Filiale 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 Zeilen für Lieferanten, EK und VPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.Label Label22 
         Caption         =   "Ausgabevarianten für Cipherlab"
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
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      MaxLength       =   6
      TabIndex        =   8
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   28
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   5175
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Top             =   7680
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   0
      Top             =   7080
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
      Caption         =   "Erstellen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   5
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      Caption         =   "Standard"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      Caption         =   "Ändern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   360
      Index           =   6
      Left            =   1560
      TabIndex        =   11
      Top             =   3720
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
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      Caption         =   "Standard"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label lbl6 
      Caption         =   "Pfad zum Converter"
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
      Index           =   2
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Y-Dateien aus der Zentrale"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lbl6 
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
      Index           =   1
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Y-Dateien aus der Zentrale"
      Top             =   3120
      Width           =   6615
   End
   Begin VB.Label lbl6 
      Caption         =   "Nur Artikel berücksichtigen, die diesem Lieferanten angehören"
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
      Index           =   0
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Y-Dateien aus der Zentrale"
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Label lbl6 
      Caption         =   "Pfad zur Artikeldatei"
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
      Index           =   92
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Y-Dateien aus der Zentrale"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "MDE - Gerät betanken"
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
      TabIndex        =   3
      Top             =   120
      Width           =   11535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL183"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FuelleZielDateiREWEMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt...", Label1(4)
    
    loeschNEW "MDE_EXPORT_REWE", gdBase
    CreateTableT2 "MDE_EXPORT_REWE", gdBase
    
    cSQL = "Insert into MDE_EXPORT_REWE Select "
    cSQL = cSQL & "'' as SCANCODE "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTIKEL "
    gdBase.Execute cSQL, dbFailOnError
    
    
    Dim rsrs            As Recordset
    Dim cSatz           As String
    
    cSQL = " Select "
    cSQL = cSQL & " SCANCODE , ARTNR  from MDE_EXPORT_REWE "

    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
            
                cSatz = rsrs!artnr
                
                rsrs.Edit
                rsrs!SCANCODE = fnMoveArtNr2EAN8(cSatz)
                rsrs.Update
                
                
            End If
            rsrs.MoveNext
        Loop

        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (ArtNr) werden erstellt...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_REWE Select "
    cSQL = cSQL & " EAN as SCANCODE"
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where Len(EAN) > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN1) werden erstellt...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_REWE Select "
    cSQL = cSQL & " EAN2 as SCANCODE "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where Len(EAN2) > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN2) werden erstellt...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_REWE Select "
    cSQL = cSQL & " EAN3 as SCANCODE "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where Len(EAN3) > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN3) werden erstellt...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_REWE Select "
    cSQL = cSQL & " ARTNR as SCANCODE "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTIKEL "
    gdBase.Execute cSQL, dbFailOnError

    'Duplikate löschen
    
    Dim cSCANCODE       As String
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lcount          As Long
    
    loeschNEW "alit" & srechnertab, gdBase
    cSQL = "select count(SCANCODE) as count ,SCANCODE into alit" & srechnertab & " from MDE_EXPORT_REWE group by SCANCODE having count(SCANCODE) > 1"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "artdupli" & srechnertab, gdBase
    cSQL = "Select * into artDupli" & srechnertab & " from MDE_EXPORT_REWE where artnr = -1 "
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsartDupli = gdBase.OpenRecordset("artDupli" & srechnertab, dbOpenTable)
    Set rsrs = gdBase.OpenRecordset("alit" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!SCANCODE) Then
                cSCANCODE = Trim(rsrs!SCANCODE)
            End If
            
            cSQL = "Select * from MDE_EXPORT_REWE where SCANCODE = '" & cSCANCODE & "'"
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst
                
                rsArt.MoveNext
                Do While Not rsArt.EOF
                    
                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).value = rsArt(i).value
                    Next i
                    rsartDupli.Update
                    
                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close: Set rsartDupli = Nothing
    
    cSQL = "Delete from MDE_EXPORT_REWE "
    cSQL = cSQL & " where val(SCANCODE) = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (VPE) werden erstellt...", Label1(4)
    
    loeschNEW "MAXVPE", gdBase
    
    cSQL = "Select artnr, max(minmen) as vpe into MAXVPE from Artlief group by artnr"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from MAXVPE where vpe < 2"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from MAXVPE where vpe is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_REWE Set "
    cSQL = cSQL & " VPE = 1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_REWE a inner join MAXVPE b on a.artnr = b.artnr Set "
    cSQL = cSQL & " a.VPE = b.vpe "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_REWE Set BEZEICH = VPE & chr(120) & space(1) & BEZEICH"
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (CSV) werden erstellt...", Label1(4)

    ExportCSV
    
    loeschNEW "alit" & srechnertab, gdBase
    loeschNEW "artdupli" & srechnertab, gdBase
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleZielDateiREWEMDE"
    Fehler.gsFehlertext = "Im Programmteil REWE-MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleZielDateiCIPHERLABMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL", gdApp
    
    anzeige "normal", "kopiere Tabelle Artlief...", Label1(4)
    
    loeschNEW "ARTLIEF", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
    
    anzeige "normal", "kopiere Tabelle Artikel...", Label1(4)
    
    loeschNEW "ARTIKEL", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTIKEL"
    
    anzeige "normal", "kopiere Tabelle Lisrt...", Label1(4)
    
    loeschNEW "LISRT", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 1...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
    cSQL = cSQL & " Artlief.Artnr "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", Artlief.LINR "
    cSQL = cSQL & ", Artlief.LEKPR "
    cSQL = cSQL & ", Artlief.MINMEN as VPE "
    cSQL = cSQL & ", '' as KUERZEL "
    cSQL = cSQL & ", '' as DEL "
    cSQL = cSQL & " from ARTLIEF "
    
    If Check6.value = vbChecked Or Check1.value = vbChecked Or Check2.value = vbChecked Or Check3.value = vbChecked Then
        cSQL = cSQL & " inner join Artikel on Artlief.artnr = artikel.artnr"
    End If
    
    cSQL = cSQL & " where Artlief.lekpr > 0 and Artlief.lekpr < 10000 "
    
    If Check6.value = vbChecked Then
        cSQL = cSQL & " and Artikel.gefuehrt = 'J'"
    End If
    
    If Check3.value = vbChecked Then
        cSQL = cSQL & " and (Artikel.bestand > 0 or Artikel.minbest > 0 )"
    Else
        If Check1.value = vbChecked Then
            cSQL = cSQL & " and Artikel.bestand > 0 "
        End If
        
        If Check2.value = vbChecked Then
            cSQL = cSQL & " and Artikel.minbest > 0 "
        End If
    End If
    
    cSQL = cSQL & " and ARTLIEF.RKZ = 'N'"
    If Text1(0).Text <> "" Then
        cSQL = cSQL & " and Artlief.LINR = " & Text1(0).Text
    End If
    gdApp.Execute cSQL, dbFailOnError
    
    If Text1(0).Text <> "" Then
    
        cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
        cSQL = cSQL & " a.Artnr "
        cSQL = cSQL & ", '' as BEZEICH "
        cSQL = cSQL & ", a.LINR "
        cSQL = cSQL & ", a.LEKPR "
        cSQL = cSQL & ", a.MINMEN as VPE "
        cSQL = cSQL & ", '' as KUERZEL "
        cSQL = cSQL & ", '' as DEL "
        cSQL = cSQL & " from ARTLIEF a inner join MDE_EXPORT_SCANPAL m on a.artnr = m.artnr where a.lekpr > 0 and a.lekpr < 10000 "
        cSQL = cSQL & " and a.LINR <> " & Text1(0).Text
        cSQL = cSQL & " and a.RKZ = 'N'"
        gdApp.Execute cSQL, dbFailOnError
        
    End If
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 2...", Label1(4)
    
    loeschNEW "temp1", gdApp
    
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp1 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp1 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp1 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp1 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on MDE_EXPORT_SCANPAL (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp1 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp1 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    '2.
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 3...", Label1(4)
    
    loeschNEW "temp2", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp2 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp2 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp2 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp2 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp2 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    '3.
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 4...", Label1(4)
    
    loeschNEW "temp3", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp3 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp3 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp3 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp3 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp3 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 5...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL2", gdApp
        
    cSQL = "Insert into MDE_EXPORT_SCANPAL2 Select "
    cSQL = cSQL & " Artnr "
    cSQL = cSQL & ", '' as SCANCODE "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", lek as LEKPR1 "
    cSQL = cSQL & ", linr as LINR1 "
    cSQL = cSQL & " from temp1 "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp2 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR2 = t.lek  "
    cSQL = cSQL & ", m.linr2 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp3 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR3 = t.lek  "
    cSQL = cSQL & ", m.linr3 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.bezeich = a.bezeich "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr1 on MDE_EXPORT_SCANPAL2 (linr1) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr2 on MDE_EXPORT_SCANPAL2 (linr2) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr3 on MDE_EXPORT_SCANPAL2 (linr3) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr1 = a.linr"
    cSQL = cSQL & " set m.MINMEN1 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr2 = a.linr"
    cSQL = cSQL & " set m.MINMEN2 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr3 = a.linr"
    cSQL = cSQL & " set m.MINMEN3 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 6...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 7...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL1 = '' where KUERZEL1 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL2 = '' where KUERZEL2 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL3 = '' where KUERZEL3 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 8...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL1 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL2 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL3 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL3", gdApp
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 9...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (ArtNr) werden erstellt...", Label1(4)
    
    Dim rsrs            As Recordset
    Dim cSatz           As String
    
    cSQL = "Select SCANCODE, ARTNR  from MDE_EXPORT_SCANPAL3 "

    Set rsrs = gdApp.OpenRecordset(cSQL)
    If Not rsrs.EOF Then

        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then

                cSatz = rsrs!artnr

                rsrs.Edit
                rsrs!SCANCODE = fnMoveArtNr2EAN8(cSatz)
                rsrs.Update
            End If
            rsrs.MoveNext
        Loop

        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing

    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN1) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean "
    cSQL = cSQL & " where Len(a.EAN) > 0 and m.scancode = '' "
    
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN2) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean2 "
    cSQL = cSQL & " where Len(a.EAN2) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN3) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean3 "
    cSQL = cSQL & " where Len(a.EAN3) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index scancode on MDE_EXPORT_SCANPAL3 (scancode) "
    gdApp.Execute cSQL, dbFailOnError


    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Duplikate löschen...", Label1(4)

    'Duplikate löschen

    Dim cSCANCODE       As String
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lcount          As Long

    loeschNEW "alit" & srechnertab, gdApp
    cSQL = "select count(SCANCODE) as count ,SCANCODE into alit" & srechnertab & " from MDE_EXPORT_SCANPAL3 group by SCANCODE having count(SCANCODE) > 1"
    gdApp.Execute cSQL, dbFailOnError

    loeschNEW "artdupli" & srechnertab, gdApp
    cSQL = "Select * into artDupli" & srechnertab & " from MDE_EXPORT_SCANPAL3 where artnr = -1 "
    gdApp.Execute cSQL, dbFailOnError

    Set rsartDupli = gdApp.OpenRecordset("artDupli" & srechnertab, dbOpenTable)
    Set rsrs = gdApp.OpenRecordset("alit" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!SCANCODE) Then
                cSCANCODE = Trim(rsrs!SCANCODE)
            End If

            cSQL = "Select * from MDE_EXPORT_SCANPAL3 where SCANCODE = '" & cSCANCODE & "'"
            Set rsArt = gdApp.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst

                rsArt.MoveNext
                Do While Not rsArt.EOF

                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).value = rsArt(i).value
                    Next i
                    rsartDupli.Update

                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close: Set rsartDupli = Nothing




    cSQL = "Delete from MDE_EXPORT_SCANPAL3 "
    cSQL = cSQL & " where val(SCANCODE) = 0 "
    gdApp.Execute cSQL, dbFailOnError
    

    
    
    

    anzeige "normal", "Artikeldaten für das MDE-Gerät (txt) werden erstellt...", Label1(4)

    ExportCSV_ScanPal

    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    
    loeschNEW "ARTLIEF", gdApp
    loeschNEW "ARTIKEL", gdApp
    loeschNEW "LISRT", gdApp
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleZielDateiCIPHERLABMDE"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleZielDateiCIPHERLABMDE_mitBestand()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL", gdApp
    
    
    If NewTableSuchenDBKombi("ZBESTAND", gdBase) = True Then
        
    
        anzeige "normal", "kopiere Tabelle ZBESTAND...", Label1(4)
        
        loeschNEW "ZBESTAND", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "ZBESTAND"
    
    End If
    
    anzeige "normal", "kopiere Tabelle Artlief...", Label1(4)
    
    loeschNEW "ARTLIEF", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
    
    anzeige "normal", "kopiere Tabelle Artikel...", Label1(4)
    
    loeschNEW "ARTIKEL", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTIKEL"
    
    anzeige "normal", "kopiere Tabelle Lisrt...", Label1(4)
    
    loeschNEW "LISRT", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 1...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
    cSQL = cSQL & " ARTLIEF.Artnr "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", ARTLIEF.LINR "
    cSQL = cSQL & ", ARTLIEF.LEKPR "
    cSQL = cSQL & ", ARTLIEF.MINMEN as VPE "
    cSQL = cSQL & ", '' as KUERZEL "
    cSQL = cSQL & ", '' as DEL "
    
    
    
    
    
    cSQL = cSQL & " from ARTLIEF "
    
    If Check6.value = vbChecked Or Check1.value = vbChecked Or Check2.value = vbChecked Or Check3.value = vbChecked Then
        cSQL = cSQL & " inner join Artikel on Artlief.artnr = artikel.artnr"
    End If
       
    
    cSQL = cSQL & " where Artlief.lekpr > 0 and Artlief.lekpr < 10000 "
    
    If Check6.value = vbChecked Then
        cSQL = cSQL & " and Artikel.gefuehrt = 'J'"
    End If
    
    If Check3.value = vbChecked Then
        cSQL = cSQL & " and (Artikel.bestand > 0 or Artikel.minbest > 0 )"
    Else
        If Check1.value = vbChecked Then
            cSQL = cSQL & " and Artikel.bestand > 0 "
        End If
        
        If Check2.value = vbChecked Then
            cSQL = cSQL & " and Artikel.minbest > 0 "
        End If
    End If
    
    cSQL = cSQL & " and ARTLIEF.RKZ = 'N'"
    If Text1(0).Text <> "" Then
        cSQL = cSQL & " and Artlief.LINR = " & Text1(0).Text
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    gdApp.Execute cSQL, dbFailOnError
    
    If Text1(0).Text <> "" Then
    
        cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
        cSQL = cSQL & " a.Artnr "
        cSQL = cSQL & ", '' as BEZEICH "
        cSQL = cSQL & ", a.LINR "
        cSQL = cSQL & ", a.LEKPR "
        cSQL = cSQL & ", a.MINMEN as VPE "
        cSQL = cSQL & ", '' as KUERZEL "
        cSQL = cSQL & ", '' as DEL "
        cSQL = cSQL & " from ARTLIEF a inner join MDE_EXPORT_SCANPAL m on a.artnr = m.artnr where a.lekpr > 0 and a.lekpr < 10000 "
        cSQL = cSQL & " and a.LINR <> " & Text1(0).Text
        cSQL = cSQL & " and a.RKZ = 'N'"
        gdApp.Execute cSQL, dbFailOnError
        
    End If
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 2...", Label1(4)
    
    
    loeschNEW "temp1", gdApp
    
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp1 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp1 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp1 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp1 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on MDE_EXPORT_SCANPAL (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp1 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp1 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    '2.
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 3...", Label1(4)
    
    loeschNEW "temp2", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp2 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp2 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp2 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp2 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp2 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    '3.
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 4...", Label1(4)
    
    loeschNEW "temp3", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp3 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp3 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp3 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp3 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp3 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 5...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL2", gdApp
        
    cSQL = "Insert into MDE_EXPORT_SCANPAL2 Select "
    cSQL = cSQL & " Artnr "
    cSQL = cSQL & ", '' as SCANCODE "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", lek as LEKPR1 "
    cSQL = cSQL & ", linr as LINR1 "
    cSQL = cSQL & " from temp1 "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp2 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR2 = t.lek  "
    cSQL = cSQL & ", m.linr2 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp3 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR3 = t.lek  "
    cSQL = cSQL & ", m.linr3 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.bezeich = a.bezeich "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr1 on MDE_EXPORT_SCANPAL2 (linr1) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr2 on MDE_EXPORT_SCANPAL2 (linr2) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr3 on MDE_EXPORT_SCANPAL2 (linr3) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr1 = a.linr"
    cSQL = cSQL & " set m.MINMEN1 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr2 = a.linr"
    cSQL = cSQL & " set m.MINMEN2 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr3 = a.linr"
    cSQL = cSQL & " set m.MINMEN3 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 6...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 7...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL1 = '' where KUERZEL1 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL2 = '' where KUERZEL2 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL3 = '' where KUERZEL3 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 8...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL1 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL2 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL3 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL3", gdApp
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 9...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (ArtNr) werden erstellt...", Label1(4)
    
    Dim rsrs            As Recordset
    Dim cSatz           As String
    
    cSQL = "Select SCANCODE, ARTNR  from MDE_EXPORT_SCANPAL3 "

    Set rsrs = gdApp.OpenRecordset(cSQL)
    If Not rsrs.EOF Then

        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then

                cSatz = rsrs!artnr

                rsrs.Edit
                rsrs!SCANCODE = fnMoveArtNr2EAN8(cSatz)
                rsrs.Update
            End If
            rsrs.MoveNext
        Loop

        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing

    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN1) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean "
    cSQL = cSQL & " where Len(a.EAN) > 0 and m.scancode = '' "
    
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN2) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean2 "
    cSQL = cSQL & " where Len(a.EAN2) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN3) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean3 "
    cSQL = cSQL & " where Len(a.EAN3) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index scancode on MDE_EXPORT_SCANPAL3 (scancode) "
    gdApp.Execute cSQL, dbFailOnError


    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Duplikate löschen...", Label1(4)

    'Duplikate löschen

    Dim cSCANCODE       As String
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lcount          As Long

    loeschNEW "alit" & srechnertab, gdApp
    cSQL = "select count(SCANCODE) as count ,SCANCODE into alit" & srechnertab & " from MDE_EXPORT_SCANPAL3 group by SCANCODE having count(SCANCODE) > 1"
    gdApp.Execute cSQL, dbFailOnError

    loeschNEW "artdupli" & srechnertab, gdApp
    cSQL = "Select * into artDupli" & srechnertab & " from MDE_EXPORT_SCANPAL3 where artnr = -1 "
    gdApp.Execute cSQL, dbFailOnError

    Set rsartDupli = gdApp.OpenRecordset("artDupli" & srechnertab, dbOpenTable)
    Set rsrs = gdApp.OpenRecordset("alit" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!SCANCODE) Then
                cSCANCODE = Trim(rsrs!SCANCODE)
            End If

            cSQL = "Select * from MDE_EXPORT_SCANPAL3 where SCANCODE = '" & cSCANCODE & "'"
            Set rsArt = gdApp.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst

                rsArt.MoveNext
                Do While Not rsArt.EOF

                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).value = rsArt(i).value
                    Next i
                    rsartDupli.Update

                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close: Set rsartDupli = Nothing




    cSQL = "Delete from MDE_EXPORT_SCANPAL3 "
    cSQL = cSQL & " where val(SCANCODE) = 0 "
    gdApp.Execute cSQL, dbFailOnError

    anzeige "normal", "Artikeldaten für das MDE-Gerät (txt) werden erstellt...", Label1(4)

    ExportCSV_ScanPal_mitBestand

    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    
    loeschNEW "ARTLIEF", gdApp
    loeschNEW "ARTIKEL", gdApp
    loeschNEW "LISRT", gdApp
    loeschNEW "ZBESTAND", gdApp
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleZielDateiCIPHERLABMDE_mitBestand"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleZielDateiCIPHERLABMDE_mitBestand_OnlyFil()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL", gdApp
    
    
    If NewTableSuchenDBKombi("Kassjour", gdBase) = True Then


        anzeige "normal", "kopiere Tabelle Kassjour...", Label1(4)

        loeschNEW "Kassjour", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "Kassjour"

    End If
    
    anzeige "normal", "kopiere Tabelle Artlief...", Label1(4)
    
    loeschNEW "ARTLIEF", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
    
    anzeige "normal", "kopiere Tabelle Artikel...", Label1(4)
    
    loeschNEW "ARTIKEL", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTIKEL"
    
    anzeige "normal", "kopiere Tabelle Lisrt...", Label1(4)
    
    loeschNEW "LISRT", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 1...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
    cSQL = cSQL & " Artlief.Artnr "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", Artlief.LINR "
    cSQL = cSQL & ", Artlief.LEKPR "
    cSQL = cSQL & ", Artlief.MINMEN as VPE "
    cSQL = cSQL & ", '' as KUERZEL "
    cSQL = cSQL & ", '' as DEL "
    
    
    
    
    
    cSQL = cSQL & " from ARTLIEF "
    
    If Check6.value = vbChecked Or Check1.value = vbChecked Or Check2.value = vbChecked Or Check3.value = vbChecked Then
        cSQL = cSQL & " inner join Artikel on Artlief.artnr = artikel.artnr"
    End If
    
    cSQL = cSQL & " where Artlief.lekpr > 0 and Artlief.lekpr < 10000 "
    
    If Check6.value = vbChecked Then
        cSQL = cSQL & " and Artikel.gefuehrt = 'J'"
    End If
    
    If Check3.value = vbChecked Then
        cSQL = cSQL & " and (Artikel.bestand > 0 or Artikel.minbest > 0 )"
    Else
        If Check1.value = vbChecked Then
            cSQL = cSQL & " and Artikel.bestand > 0 "
        End If
        
        If Check2.value = vbChecked Then
            cSQL = cSQL & " and Artikel.minbest > 0 "
        End If
    End If
    
    cSQL = cSQL & " and ARTLIEF.RKZ = 'N'"
    If Text1(0).Text <> "" Then
        cSQL = cSQL & " and Artlief.LINR = " & Text1(0).Text
    End If
    
    
    
    
    
    
    
    
    
    gdApp.Execute cSQL, dbFailOnError
    
    If Text1(0).Text <> "" Then
    
        cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
        cSQL = cSQL & " a.Artnr "
        cSQL = cSQL & ", '' as BEZEICH "
        cSQL = cSQL & ", a.LINR "
        cSQL = cSQL & ", a.LEKPR "
        cSQL = cSQL & ", a.MINMEN as VPE "
        cSQL = cSQL & ", '' as KUERZEL "
        cSQL = cSQL & ", '' as DEL "
        cSQL = cSQL & " from ARTLIEF a inner join MDE_EXPORT_SCANPAL m on a.artnr = m.artnr where a.lekpr > 0 and a.lekpr < 10000 "
        cSQL = cSQL & " and a.LINR <> " & Text1(0).Text
        cSQL = cSQL & " and a.RKZ = 'N'"
        gdApp.Execute cSQL, dbFailOnError
        
    End If
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 2...", Label1(4)
    
    
    loeschNEW "temp1", gdApp
    
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp1 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp1 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp1 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp1 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on MDE_EXPORT_SCANPAL (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp1 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp1 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    '2.
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 3...", Label1(4)
    
    loeschNEW "temp2", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp2 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp2 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp2 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp2 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp2 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    '3.
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 4...", Label1(4)
    
    loeschNEW "temp3", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp3 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp3 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp3 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp3 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp3 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 5...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL2", gdApp
        
    cSQL = "Insert into MDE_EXPORT_SCANPAL2 Select "
    cSQL = cSQL & " Artnr "
    cSQL = cSQL & ", '' as SCANCODE "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", lek as LEKPR1 "
    cSQL = cSQL & ", linr as LINR1 "
    cSQL = cSQL & " from temp1 "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp2 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR2 = t.lek  "
    cSQL = cSQL & ", m.linr2 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp3 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR3 = t.lek  "
    cSQL = cSQL & ", m.linr3 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.bezeich = a.bezeich "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr1 on MDE_EXPORT_SCANPAL2 (linr1) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr2 on MDE_EXPORT_SCANPAL2 (linr2) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr3 on MDE_EXPORT_SCANPAL2 (linr3) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr1 = a.linr"
    cSQL = cSQL & " set m.MINMEN1 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr2 = a.linr"
    cSQL = cSQL & " set m.MINMEN2 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr3 = a.linr"
    cSQL = cSQL & " set m.MINMEN3 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 6...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 7...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL1 = '' where KUERZEL1 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL2 = '' where KUERZEL2 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL3 = '' where KUERZEL3 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 8...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL1 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL2 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL3 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL3", gdApp
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 9...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (ArtNr) werden erstellt...", Label1(4)
    
    Dim rsrs            As Recordset
    Dim cSatz           As String
    
    cSQL = "Select SCANCODE, ARTNR  from MDE_EXPORT_SCANPAL3 "

    Set rsrs = gdApp.OpenRecordset(cSQL)
    If Not rsrs.EOF Then

        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then

                cSatz = rsrs!artnr

                rsrs.Edit
                rsrs!SCANCODE = fnMoveArtNr2EAN8(cSatz)
                rsrs.Update
            End If
            rsrs.MoveNext
        Loop

        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing

    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN1) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean "
    cSQL = cSQL & " where Len(a.EAN) > 0 and m.scancode = '' "
    
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN2) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean2 "
    cSQL = cSQL & " where Len(a.EAN2) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN3) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean3 "
    cSQL = cSQL & " where Len(a.EAN3) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index scancode on MDE_EXPORT_SCANPAL3 (scancode) "
    gdApp.Execute cSQL, dbFailOnError


    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Duplikate löschen...", Label1(4)

    'Duplikate löschen

    Dim cSCANCODE       As String
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lcount          As Long

    loeschNEW "alit" & srechnertab, gdApp
    cSQL = "select count(SCANCODE) as count ,SCANCODE into alit" & srechnertab & " from MDE_EXPORT_SCANPAL3 group by SCANCODE having count(SCANCODE) > 1"
    gdApp.Execute cSQL, dbFailOnError

    loeschNEW "artdupli" & srechnertab, gdApp
    cSQL = "Select * into artDupli" & srechnertab & " from MDE_EXPORT_SCANPAL3 where artnr = -1 "
    gdApp.Execute cSQL, dbFailOnError

    Set rsartDupli = gdApp.OpenRecordset("artDupli" & srechnertab, dbOpenTable)
    Set rsrs = gdApp.OpenRecordset("alit" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!SCANCODE) Then
                cSCANCODE = Trim(rsrs!SCANCODE)
            End If

            cSQL = "Select * from MDE_EXPORT_SCANPAL3 where SCANCODE = '" & cSCANCODE & "'"
            Set rsArt = gdApp.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst

                rsArt.MoveNext
                Do While Not rsArt.EOF

                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).value = rsArt(i).value
                    Next i
                    rsartDupli.Update

                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close: Set rsartDupli = Nothing




    cSQL = "Delete from MDE_EXPORT_SCANPAL3 "
    cSQL = cSQL & " where val(SCANCODE) = 0 "
    gdApp.Execute cSQL, dbFailOnError

    anzeige "normal", "Artikeldaten für das MDE-Gerät (txt) werden erstellt...", Label1(4)

    ExportCSV_ScanPal_mitBestand_OnlyFil

    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    
    loeschNEW "ARTLIEF", gdApp
    loeschNEW "ARTIKEL", gdApp
    loeschNEW "LISRT", gdApp
    loeschNEW "Kassjour", gdApp
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleZielDateiCIPHERLABMDE_mitBestand_OnlyFil"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleZielDateiCIPHERLABMDE_mitBestandKVK_OnlyFil()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL", gdApp
    
    
    If NewTableSuchenDBKombi("Kassjour", gdBase) = True Then


        anzeige "normal", "kopiere Tabelle Kassjour...", Label1(4)

        loeschNEW "Kassjour", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "Kassjour"

    End If
    
    anzeige "normal", "kopiere Tabelle Artlief...", Label1(4)
    
    loeschNEW "ARTLIEF", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
    
    anzeige "normal", "kopiere Tabelle Artikel...", Label1(4)
    
    loeschNEW "ARTIKEL", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTIKEL"
    
    anzeige "normal", "kopiere Tabelle Lisrt...", Label1(4)
    
    loeschNEW "LISRT", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 1...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
    cSQL = cSQL & " Artlief.Artnr "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", Artlief.LINR "
    cSQL = cSQL & ", Artlief.LEKPR "
    cSQL = cSQL & ", Artlief.MINMEN as VPE "
    cSQL = cSQL & ", '' as KUERZEL "
    cSQL = cSQL & ", '' as DEL "
    
    
    
    
    
    cSQL = cSQL & " from ARTLIEF "
    
    If Check6.value = vbChecked Or Check1.value = vbChecked Or Check2.value = vbChecked Or Check3.value = vbChecked Then
        cSQL = cSQL & " inner join Artikel on Artlief.artnr = artikel.artnr"
    End If
    
    cSQL = cSQL & " where Artlief.lekpr > 0 and Artlief.lekpr < 10000 "
    
    If Check6.value = vbChecked Then
        cSQL = cSQL & " and Artikel.gefuehrt = 'J'"
    End If
    
    If Check3.value = vbChecked Then
        cSQL = cSQL & " and (Artikel.bestand > 0 or Artikel.minbest > 0 )"
    Else
        If Check1.value = vbChecked Then
            cSQL = cSQL & " and Artikel.bestand > 0 "
        End If
        
        If Check2.value = vbChecked Then
            cSQL = cSQL & " and Artikel.minbest > 0 "
        End If
    End If
    
    cSQL = cSQL & " and ARTLIEF.RKZ = 'N'"
    If Text1(0).Text <> "" Then
        cSQL = cSQL & " and Artlief.LINR = " & Text1(0).Text
    End If
    
    
    
    
    
    
    
    
    
    gdApp.Execute cSQL, dbFailOnError
    
    If Text1(0).Text <> "" Then
    
        cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
        cSQL = cSQL & " a.Artnr "
        cSQL = cSQL & ", '' as BEZEICH "
        cSQL = cSQL & ", a.LINR "
        cSQL = cSQL & ", a.LEKPR "
        cSQL = cSQL & ", a.MINMEN as VPE "
        cSQL = cSQL & ", '' as KUERZEL "
        cSQL = cSQL & ", '' as DEL "
        cSQL = cSQL & " from ARTLIEF a inner join MDE_EXPORT_SCANPAL m on a.artnr = m.artnr where a.lekpr > 0 and a.lekpr < 10000 "
        cSQL = cSQL & " and a.LINR <> " & Text1(0).Text
        cSQL = cSQL & " and a.RKZ = 'N'"
        gdApp.Execute cSQL, dbFailOnError
        
    End If
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 2...", Label1(4)
    
    
    loeschNEW "temp1", gdApp
    
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp1 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp1 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp1 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp1 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on MDE_EXPORT_SCANPAL (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp1 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp1 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    '2.
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 3...", Label1(4)
    
    loeschNEW "temp2", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp2 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp2 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp2 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp2 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp2 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    '3.
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 4...", Label1(4)
    
    loeschNEW "temp3", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp3 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp3 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp3 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp3 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp3 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 5...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL2", gdApp
        
    cSQL = "Insert into MDE_EXPORT_SCANPAL2 Select "
    cSQL = cSQL & " Artnr "
    cSQL = cSQL & ", '' as SCANCODE "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", lek as LEKPR1 "
    cSQL = cSQL & ", linr as LINR1 "
    cSQL = cSQL & " from temp1 "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp2 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR2 = t.lek  "
    cSQL = cSQL & ", m.linr2 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp3 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR3 = t.lek  "
    cSQL = cSQL & ", m.linr3 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.bezeich = a.bezeich "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr1 on MDE_EXPORT_SCANPAL2 (linr1) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr2 on MDE_EXPORT_SCANPAL2 (linr2) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr3 on MDE_EXPORT_SCANPAL2 (linr3) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr1 = a.linr"
    cSQL = cSQL & " set m.MINMEN1 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr2 = a.linr"
    cSQL = cSQL & " set m.MINMEN2 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr3 = a.linr"
    cSQL = cSQL & " set m.MINMEN3 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 6...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 7...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL1 = '' where KUERZEL1 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL2 = '' where KUERZEL2 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL3 = '' where KUERZEL3 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 8...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL1 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL2 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL3 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL3", gdApp
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 9...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (ArtNr) werden erstellt...", Label1(4)
    
    Dim rsrs            As Recordset
    Dim cSatz           As String
    
    cSQL = "Select SCANCODE, ARTNR  from MDE_EXPORT_SCANPAL3 "

    Set rsrs = gdApp.OpenRecordset(cSQL)
    If Not rsrs.EOF Then

        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then

                cSatz = rsrs!artnr

                rsrs.Edit
                rsrs!SCANCODE = fnMoveArtNr2EAN8(cSatz)
                rsrs.Update
            End If
            rsrs.MoveNext
        Loop

        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing

    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN1) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean "
    cSQL = cSQL & " where Len(a.EAN) > 0 and m.scancode = '' "
    
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN2) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean2 "
    cSQL = cSQL & " where Len(a.EAN2) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN3) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean3 "
    cSQL = cSQL & " where Len(a.EAN3) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index scancode on MDE_EXPORT_SCANPAL3 (scancode) "
    gdApp.Execute cSQL, dbFailOnError


    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Duplikate löschen...", Label1(4)

    'Duplikate löschen

    Dim cSCANCODE       As String
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lcount          As Long

    loeschNEW "alit" & srechnertab, gdApp
    cSQL = "select count(SCANCODE) as count ,SCANCODE into alit" & srechnertab & " from MDE_EXPORT_SCANPAL3 group by SCANCODE having count(SCANCODE) > 1"
    gdApp.Execute cSQL, dbFailOnError

    loeschNEW "artdupli" & srechnertab, gdApp
    cSQL = "Select * into artDupli" & srechnertab & " from MDE_EXPORT_SCANPAL3 where artnr = -1 "
    gdApp.Execute cSQL, dbFailOnError

    Set rsartDupli = gdApp.OpenRecordset("artDupli" & srechnertab, dbOpenTable)
    Set rsrs = gdApp.OpenRecordset("alit" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!SCANCODE) Then
                cSCANCODE = Trim(rsrs!SCANCODE)
            End If

            cSQL = "Select * from MDE_EXPORT_SCANPAL3 where SCANCODE = '" & cSCANCODE & "'"
            Set rsArt = gdApp.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst

                rsArt.MoveNext
                Do While Not rsArt.EOF

                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).value = rsArt(i).value
                    Next i
                    rsartDupli.Update

                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close: Set rsartDupli = Nothing

    cSQL = "Delete from MDE_EXPORT_SCANPAL3 "
    cSQL = cSQL & " where val(SCANCODE) = 0 "
    gdApp.Execute cSQL, dbFailOnError

    anzeige "normal", "Artikeldaten für das MDE-Gerät (txt) werden erstellt...", Label1(4)

    ExportCSV_ScanPal_mitBestandKVK_OnlyFil

    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    
    loeschNEW "ARTLIEF", gdApp
    loeschNEW "ARTIKEL", gdApp
    loeschNEW "LISRT", gdApp
    loeschNEW "Kassjour", gdApp
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleZielDateiCIPHERLABMDE_mitBestandKVK_OnlyFil"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub





Private Sub FuelleZielDateiCIPHERLABMDE_mitKVK_OnlyFil()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL", gdApp
    
    
    If NewTableSuchenDBKombi("Zugang", gdBase) = True Then


        anzeige "normal", "kopiere Tabelle Zugang...", Label1(4)

        loeschNEW "Zugang", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "Zugang"

    End If
    
    anzeige "normal", "kopiere Tabelle Artlief...", Label1(4)
    
    loeschNEW "ARTLIEF", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
    
    anzeige "normal", "kopiere Tabelle Artikel...", Label1(4)
    
    loeschNEW "ARTIKEL", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTIKEL"
    
    anzeige "normal", "kopiere Tabelle Lisrt...", Label1(4)
    
    loeschNEW "LISRT", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 1...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
    cSQL = cSQL & " Artlief.Artnr "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", Artlief.LINR "
    cSQL = cSQL & ", Artlief.LEKPR "
    cSQL = cSQL & ", Artlief.MINMEN as VPE "
    cSQL = cSQL & ", '' as KUERZEL "
    cSQL = cSQL & ", '' as DEL "

    cSQL = cSQL & " from ARTLIEF "
    
    If Check6.value = vbChecked Or Check1.value = vbChecked Or Check2.value = vbChecked Or Check3.value = vbChecked Then
        cSQL = cSQL & " inner join Artikel on Artlief.artnr = artikel.artnr"
    End If
    
    cSQL = cSQL & " where Artlief.lekpr > 0 and Artlief.lekpr < 10000 "
    
    If Check6.value = vbChecked Then
        cSQL = cSQL & " and Artikel.gefuehrt = 'J'"
    End If
    
    If Check3.value = vbChecked Then
        cSQL = cSQL & " and (Artikel.bestand > 0 or Artikel.minbest > 0 )"
    Else
        If Check1.value = vbChecked Then
            cSQL = cSQL & " and Artikel.bestand > 0 "
        End If
        
        If Check2.value = vbChecked Then
            cSQL = cSQL & " and Artikel.minbest > 0 "
        End If
    End If
    
    cSQL = cSQL & " and ARTLIEF.RKZ = 'N'"
    If Text1(0).Text <> "" Then
        cSQL = cSQL & " and Artlief.LINR = " & Text1(0).Text
    End If
    
    
    
    
    
    
    
    
    
    gdApp.Execute cSQL, dbFailOnError
    
    If Text1(0).Text <> "" Then
    
        cSQL = "Insert into MDE_EXPORT_SCANPAL Select "
        cSQL = cSQL & " a.Artnr "
        cSQL = cSQL & ", '' as BEZEICH "
        cSQL = cSQL & ", a.LINR "
        cSQL = cSQL & ", a.LEKPR "
        cSQL = cSQL & ", a.MINMEN as VPE "
        cSQL = cSQL & ", '' as KUERZEL "
        cSQL = cSQL & ", '' as DEL "
        cSQL = cSQL & " from ARTLIEF a inner join MDE_EXPORT_SCANPAL m on a.artnr = m.artnr where a.lekpr > 0 and a.lekpr < 10000 "
        cSQL = cSQL & " and a.LINR <> " & Text1(0).Text
        cSQL = cSQL & " and a.RKZ = 'N'"
        gdApp.Execute cSQL, dbFailOnError
        
    End If
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 2...", Label1(4)
    
    
    loeschNEW "temp1", gdApp
    
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp1 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp1 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp1 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp1 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on MDE_EXPORT_SCANPAL (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp1 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp1 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    '2.
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 3...", Label1(4)
    
    loeschNEW "temp2", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp2 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp2 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp2 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp2 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp2 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    '3.
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 4...", Label1(4)
    
    loeschNEW "temp3", gdApp
    cSQL = "Select artnr,0 as linr ,min(lekpr) as lek into temp3 from MDE_EXPORT_SCANPAL group by artnr  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index artnr on temp3 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index lek on temp3 (lek) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on temp3 (linr) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update temp3 t inner join MDE_EXPORT_SCANPAL m on t.artnr = m.artnr and t.lek = m.lekpr "
    cSQL = cSQL & " set t.linr = m.linr "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL m inner join  temp3 t on m.artnr = t.artnr and m.linr = t.linr "
    cSQL = cSQL & " set DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL "
    cSQL = cSQL & " where DEL = 'X' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 5...", Label1(4)
    
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL2", gdApp
        
    cSQL = "Insert into MDE_EXPORT_SCANPAL2 Select "
    cSQL = cSQL & " Artnr "
    cSQL = cSQL & ", '' as SCANCODE "
    cSQL = cSQL & ", '' as BEZEICH "
    cSQL = cSQL & ", lek as LEKPR1 "
    cSQL = cSQL & ", linr as LINR1 "
    cSQL = cSQL & " from temp1 "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Create index artnr on MDE_EXPORT_SCANPAL2 (artnr) "
    gdApp.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp2 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR2 = t.lek  "
    cSQL = cSQL & ", m.linr2 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join temp3 t on m.artnr = t.artnr "
    cSQL = cSQL & " set m.LEKPR3 = t.lek  "
    cSQL = cSQL & ", m.linr3 = t.LINR "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.bezeich = a.bezeich "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr1 on MDE_EXPORT_SCANPAL2 (linr1) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr2 on MDE_EXPORT_SCANPAL2 (linr2) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr3 on MDE_EXPORT_SCANPAL2 (linr3) "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr1 = a.linr"
    cSQL = cSQL & " set m.MINMEN1 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr2 = a.linr"
    cSQL = cSQL & " set m.MINMEN2 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 m inner join artlief a on m.artnr = a.artnr and m.linr3 = a.linr"
    cSQL = cSQL & " set m.MINMEN3 = a.minmen "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 6...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = lisrt.KUERZEL "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 7...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL1 = '' where KUERZEL1 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL2 = '' where KUERZEL2 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 set KUERZEL3 = '' where KUERZEL3 is null "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 8...", Label1(4)
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr1 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL1 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL1 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr2 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL2 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL2 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL2 inner join lisrt on lisrt.linr = MDE_EXPORT_SCANPAL2.linr3 set "
    cSQL = cSQL & " MDE_EXPORT_SCANPAL2.KUERZEL3 = Ucase(left(lisrt.liefbez,5)) "
    cSQL = cSQL & " where MDE_EXPORT_SCANPAL2.KUERZEL3 = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    
    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    CreateTableT2 "MDE_EXPORT_SCANPAL3", gdApp
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Schritt 9...", Label1(4)
    
    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (ArtNr) werden erstellt...", Label1(4)
    
    Dim rsrs            As Recordset
    Dim cSatz           As String
    
    cSQL = "Select SCANCODE, ARTNR  from MDE_EXPORT_SCANPAL3 "

    Set rsrs = gdApp.OpenRecordset(cSQL)
    If Not rsrs.EOF Then

        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then

                cSatz = rsrs!artnr

                rsrs.Edit
                rsrs!SCANCODE = fnMoveArtNr2EAN8(cSatz)
                rsrs.Update
            End If
            rsrs.MoveNext
        Loop

        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing

    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN1) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean "
    cSQL = cSQL & " where Len(a.EAN) > 0 and m.scancode = '' "
    
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN2) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean2 "
    cSQL = cSQL & " where Len(a.EAN2) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikeldaten für das MDE-Gerät (EAN3) werden erstellt...", Label1(4)

    cSQL = "Insert into MDE_EXPORT_SCANPAL3 Select * "
    cSQL = cSQL & " from MDE_EXPORT_SCANPAL2 "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT_SCANPAL3 m inner join artikel a on m.artnr = a.artnr "
    cSQL = cSQL & " set m.scancode = a.ean3 "
    cSQL = cSQL & " where Len(a.EAN3) > 0 and m.scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from MDE_EXPORT_SCANPAL3 where scancode = '' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Create index scancode on MDE_EXPORT_SCANPAL3 (scancode) "
    gdApp.Execute cSQL, dbFailOnError


    anzeige "normal", "Artikeldaten für das MDE-Gerät werden erstellt, Duplikate löschen...", Label1(4)

    'Duplikate löschen

    Dim cSCANCODE       As String
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lcount          As Long

    loeschNEW "alit" & srechnertab, gdApp
    cSQL = "select count(SCANCODE) as count ,SCANCODE into alit" & srechnertab & " from MDE_EXPORT_SCANPAL3 group by SCANCODE having count(SCANCODE) > 1"
    gdApp.Execute cSQL, dbFailOnError

    loeschNEW "artdupli" & srechnertab, gdApp
    cSQL = "Select * into artDupli" & srechnertab & " from MDE_EXPORT_SCANPAL3 where artnr = -1 "
    gdApp.Execute cSQL, dbFailOnError

    Set rsartDupli = gdApp.OpenRecordset("artDupli" & srechnertab, dbOpenTable)
    Set rsrs = gdApp.OpenRecordset("alit" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!SCANCODE) Then
                cSCANCODE = Trim(rsrs!SCANCODE)
            End If

            cSQL = "Select * from MDE_EXPORT_SCANPAL3 where SCANCODE = '" & cSCANCODE & "'"
            Set rsArt = gdApp.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst

                rsArt.MoveNext
                Do While Not rsArt.EOF

                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).value = rsArt(i).value
                    Next i
                    rsartDupli.Update

                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close: Set rsartDupli = Nothing




    cSQL = "Delete from MDE_EXPORT_SCANPAL3 "
    cSQL = cSQL & " where val(SCANCODE) = 0 "
    gdApp.Execute cSQL, dbFailOnError

    anzeige "normal", "Artikeldaten für das MDE-Gerät (txt) werden erstellt...", Label1(4)

    
    ExportCSV_ScanPal_mitKVK_OnlyFil

    loeschNEW "MDE_EXPORT_SCANPAL3", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL2", gdApp
    loeschNEW "MDE_EXPORT_SCANPAL", gdApp
    
    loeschNEW "ARTLIEF", gdApp
    loeschNEW "ARTIKEL", gdApp
    loeschNEW "LISRT", gdApp
    loeschNEW "Zugang", gdApp
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleZielDateiCIPHERLABMDE_mitKVK_OnlyFil"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherTankpfad()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset

    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!TankPfad = gsTankPfad
        rsrs!ConverterPfad = gsConverterPfad
        rsrs.Update
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "speicherTankpfad"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        

End Sub



Private Sub Check3_Click()
On Error GoTo LOKAL_ERROR

If Check3.value = vbChecked Then
    Check2.Visible = False
    Check1.Visible = False
    Check2.value = vbUnchecked
    Check1.value = vbUnchecked
Else
    Check2.Visible = True
    Check1.Visible = True
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check3_Click"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    
    Select Case Index
    
        Case Is = 0 'Ändern
            sTitle = "Pfad zur Artikeldatei"
             sFilter = "exe - Dateien (*.exe)| *.exe"
            
            sOldpfad = "C:\"
            gsTankPfad = pfadaendern(sTitle, sFilter, sOldpfad)

   
            Text1(28).Text = gsTankPfad
            
        Case Is = 5 'Standard

            
            Text1(28).Text = "C:\Betanken"
            gsTankPfad = "C:\Betanken"
            
        Case Is = 1 'Standard

            
            Text1(1).Text = "C:\Program Files (x86)\CipherLab\Data Converter 3"
            gsConverterPfad = "C:\Program Files (x86)\CipherLab\Data Converter 3"
            
        Case 6
            Text1_KeyUp 0, vbKeyF2, 0
    End Select
    
    speicherTankpfad
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
 On Error GoTo LOKAL_ERROR
    Dim i As Integer
    Dim iFileNr As Integer
    Screen.MousePointer = 11
    
    Select Case Index
        Case 0
            gsTankPfad = Text1(28).Text
            gsConverterPfad = Text1(1).Text
            speicherTankpfad
            voreinstellungspeichernE183
            Unload frmWKL183
        Case 1      'Ziel-Datei füllen
            If gsMDEGERAET = "REWEMDE" Then
                FuelleZielDateiREWEMDE
            ElseIf gsMDEGERAET = "CIPHERLAB" Then
                gsTankPfad = Text1(28).Text
                gsConverterPfad = Text1(1).Text
                speicherTankpfad
                
                If Option1(0).value = True Then
                    FuelleZielDateiCIPHERLABMDE
                ElseIf Option1(1).value = True Then
                    FuelleZielDateiCIPHERLABMDE_mitBestand
                ElseIf Option1(2).value = True Then
                    FuelleZielDateiCIPHERLABMDE_mitBestand_OnlyFil
                ElseIf Option1(3).value = True Then
                    FuelleZielDateiCIPHERLABMDE_mitKVK_OnlyFil
                ElseIf Option1(4).value = True Then
                    FuelleZielDateiCIPHERLABMDE_mitBestandKVK_OnlyFil
                End If
                
                
            End If
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    Dim bMDECONVfound As Boolean
    Dim sMDECPfad As String
    Dim sMDECname As String
    
    Dim sLFNR As String
    
    If gcFilNr = "0" Then
        Option1(1).Visible = False
    End If
    
    If NewTableSuchenDBKombi("E183", gdApp) Then
        voreinstellungladenE183
    End If
    
    
    sLFNR = "0"
    
    If gsMDEGERAET = "CIPHERLAB" Then
        Frame7.Visible = True
    Else
        Frame7.Visible = False
    End If
    
    Text1(28).Text = gsTankPfad
    Text1(1).Text = gsConverterPfad
    
    If gsConverterPfad = "" Then
        MsgBox "Bitte geben Sie den Pfad zum Converter an!", vbCritical, "Winkiss Information:"
        
        Exit Sub
    End If
    
    sLFNR = "1"
    
    sMDECPfad = Text1(1).Text '"C:\Program Files (x86)\CipherLab\Data Converter 3"
    sMDECname = "Converter.exe"
        
    bMDECONVfound = False
    
    'close anwendung
    Dim hwnd&
    Dim Y As String
    Dim result&
    Dim Title$

    Y = "Cipher"

    lRet = Shell("taskkill /F /IM Converter.exe")

    sLFNR = "2"
    Pause 5

    hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)

    Do
        result = GetWindowTextLength(hwnd) + 1
        Title = Space(result)
        result = GetWindowText(hwnd, Title, result)
        Title = Left$(Title, Len(Title) - 1)

        If InStr(1, Title, Y) Then
            bMDECONVfound = True
            Exit Do
        End If

        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop Until hwnd = 0

    sLFNR = "3"

    If bMDECONVfound = False Then
        'Starte anwendung
        Dim prev_dir As String
        ' Save the current directory.
        prev_dir = CurDir
        ' Go to the desired startup directory.

        sLFNR = "31"

        ChDrive Left(sMDECPfad, 1)

        sLFNR = "32"
        ChDir sMDECPfad
        ' Shell the application.

        sLFNR = "33"
        lRet = Shell(sMDECname, vbNormalFocus)
        ' Restore the saved directory.

        sLFNR = "34"
        ChDir prev_dir
    End If

    sLFNR = "4"
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten. " & sLFNR
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE183()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    
    Dim bo0     As Integer
    Dim bo1     As Integer
    Dim bo2     As Integer
    Dim bo3     As Integer
    Dim bo4     As Integer
    
    loeschNEW "E183", gdApp
    CreateTableT3 "E183", gdApp
    
    bo0 = Option1(0).value
    bo1 = Option1(1).value
    bo2 = Option1(2).value
    bo3 = Option1(3).value
    bo4 = Option1(4).value
    
    sSQL = "Insert into E183 (bo0,bo1,bo2,bo3,bo4) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & "," & bo2 & "," & bo3
    sSQL = sSQL & "," & bo4 & ")"
    gdApp.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE183"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE183()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdApp.OpenRecordset("E183")
    If Not rs.EOF Then
        Option1(0).value = rs!bo0
        Option1(1).value = rs!bo1
        Option1(2).value = rs!bo2
        Option1(3).value = rs!bo3
        Option1(4).value = rs!bo4
    End If
    rs.Close: Set rs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE183"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub ExportCSV()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer
    Dim cPreis          As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    sSQL = " Select "
    sSQL = sSQL & " SCANCODE  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", KVKPR1  "
    sSQL = sSQL & " from MDE_EXPORT_REWE order by Artnr "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        sAusgabedatname = "ARTIKEL.CSV"

        cPfad1 = "C:\MDE\STAMMDAT\"
        

        cdatei = cPfad1 & sAusgabedatname
        cPfad = cPfad1
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
'        cSatz = "EAN;BEZEICH;KVKPR" & Chr$(13) & Chr$(10)

'        lPos = LOF(iFileNr)
'        lPos = lPos + 1
'        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 2
                If Not IsNull(rsrs.Fields(i)) Then

                    If i > 0 Then
                        If i = 2 Then
                            If rsrs.Fields(i) = 0 Then
                                cSatz = cSatz & ";0.00"
                            Else
                                cPreis = Format(rsrs.Fields(i), "###0.00")
                                cPreis = SwapStr(cPreis, ",", ".")
                                cSatz = cSatz & ";" & cPreis
                            End If
                        Else
                            cSatz = cSatz & ";" & rsrs.Fields(i)
                        End If
                    Else
                        cSatz = rsrs.Fields(i)
                    End If
                Else
                    If i > 0 Then
                        cSatz = cSatz & ";"
                    Else
                        cSatz = ""
                    End If
                End If
            Next i
        
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("MDE_EXPORT_REWE", gdBase) Then
        
        MsgBox "Die Dateiausgabe war erfolgreich! Führen Sie die Weitervearbeitung laut Anleitung direkt am MDE - Gerät durch!", vbInformation, "Winkiss Information:"
        
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportCSV"
        Fehler.gsFehlertext = "Im Programmteil REWE-MDE betanken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ExportCSV_ScanPal()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer
    Dim cPreis          As String
    Dim cFeld           As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    If gsTankPfad <> "" Then
        cPfad1 = gsTankPfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
    Else
        cPfad1 = gcDBPfad      'dbpfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
        cPfad1 = cPfad1 & "Box\"
    End If
    
    sSQL = "Select "
    sSQL = sSQL & " SCANCODE  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", KUERZEL1 "
    sSQL = sSQL & ", MINMEN1 "
    sSQL = sSQL & ", LEKPR1 "
    
    sSQL = sSQL & ", KUERZEL2 "
    sSQL = sSQL & ", MINMEN2 "
    sSQL = sSQL & ", LEKPR2 "
    
    sSQL = sSQL & ", KUERZEL3 "
    sSQL = sSQL & ", MINMEN3 "
    sSQL = sSQL & ", LEKPR3 "

    sSQL = sSQL & " from MDE_EXPORT_SCANPAL3 order by Artnr "
    
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        sAusgabedatname = "ARTIKEL_for_MDE.txt"
        
        cdatei = cPfad1 & sAusgabedatname
        cPfad = cPfad1
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            cFeld = ""
            
            If Not IsNull(rsrs!SCANCODE) Then
                cFeld = Left(rsrs!SCANCODE, 13)
            End If
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(13 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = Left(rsrs!BEZEICH, 35)
            End If
            
            cFeld = SwapStr(cFeld, ",", ".")
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(35 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL1) Then
                cFeld = Left(rsrs!KUERZEL1, 5)
            End If
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(7 - Len(cFeld))
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR1) Then
                cFeld = Format(rsrs!LEKPR1, "###0.00")
            End If
            
            
            cFeld = SwapStr(cFeld, ",", ".")
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(9 - Len(cFeld))
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN1) Then
                cFeld = Left(rsrs!MINMEN1, 4)
            End If
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(4 - Len(cFeld)) & ","
            
            '2.
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL2) Then
                cFeld = Left(rsrs!KUERZEL2, 5)
            End If
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(7 - Len(cFeld))
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR2) Then
                cFeld = Format(rsrs!LEKPR2, "###0.00")
            End If
            
            
            cFeld = SwapStr(cFeld, ",", ".")
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(9 - Len(cFeld))
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN2) Then
                cFeld = Left(rsrs!MINMEN2, 4)
            End If
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(4 - Len(cFeld)) & ","
            
            '3.
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL3) Then
                cFeld = Left(rsrs!KUERZEL3, 5)
            End If
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(7 - Len(cFeld))
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR3) Then
                cFeld = Format(rsrs!LEKPR3, "###0.00")
            End If
            
            
            cFeld = SwapStr(cFeld, ",", ".")
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(9 - Len(cFeld))
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN3) Then
                cFeld = Left(rsrs!MINMEN3, 4)
            End If
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(4 - Len(cFeld))
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("MDE_EXPORT_SCANPAL3", gdApp) Then
    
        Dim lTimeOut As Long
        
        lTimeOut = 200
        
        Dim iErr_Zaehler As Integer
        Dim iTimeout As Integer
        
        Dim bTimeout_erreicht As Boolean
        Dim ctmp As String
        
        bTimeout_erreicht = False
        
        iTimeout = lTimeOut * 10
        
        iErr_Zaehler = 0
        Do While FileExists(cdatei) = True
        
            anzeige "rot2", "MDE-Gerät wird betankt, bitte warten(" & iErr_Zaehler & ")...", Label1(4)
            PauseSi 0.1
            iErr_Zaehler = iErr_Zaehler + 1
            If iErr_Zaehler > iTimeout Then
                bTimeout_erreicht = True
                Exit Do
            End If
            
            
        Loop
        
        If bTimeout_erreicht = True Then
'            ctmp = "Fehler - Zeitlimit erreicht" & vbCrLf & vbCrLf

            ctmp = "Das hat leider nicht geklappt." & vbCrLf & vbCrLf
            ctmp = ctmp & "'SD Card is ready' - auf dem Display?" & vbCrLf
            ctmp = ctmp & "MDE Gerät richtig eingesteckt?" & vbCrLf & vbCrLf
            ctmp = ctmp & "Diesen Programmteil bitte über 'Schließen' verlassen und neubetreten!"
            
            iRet = MsgBox(ctmp, vbCritical + vbOKOnly, "Winkiss Hinweis")
            
            Exit Sub
        Else
        
            Pause 2
            anzeige "ERFOLG", "Das MDE-Gerät ist jetzt erfolgreich betankt.", Label1(4)
            MsgBox "Das MDE-Gerät ist jetzt erfolgreich betankt.", vbInformation, "Winkiss Hinweis:"
        
'            MsgBox "Die Dateiausgabe war erfolgreich! Die Datei: 'ARTIKEL_for_MDE.txt' befindet sich unter: " & cPfad1, vbInformation, "Winkiss Information:"
            
        End If
        
        

        
        
        
        
'        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportCSV_ScanPal"
        Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub ExportCSV_ScanPal_mitBestand()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer
    Dim cPreis          As String
    Dim cFeld           As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    If gsTankPfad <> "" Then
        cPfad1 = gsTankPfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
    Else
        cPfad1 = gcDBPfad      'dbpfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
        cPfad1 = cPfad1 & "Box\"
    End If
    
    
    loeschNEW "MDE_EXPORT_ZBESTAND", gdApp
    
    
    
    sSQL = "Select " & gcFilNr & " as Filialnr, artnr,bestand,minbest  into MDE_EXPORT_ZBESTAND from Artikel  where artnr in (Select Artnr from MDE_EXPORT_SCANPAL3)"
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    
    sSQL = "Insert into MDE_EXPORT_ZBESTAND "
    
    If gcFilNr = "1" Then
        sSQL = sSQL & " Select 2 as Filialnr "
    ElseIf gcFilNr = "2" Then
        sSQL = sSQL & " Select 1 as Filialnr "
    End If
    sSQL = sSQL & ", artnr,0 as bestand,0 as minbest  from MDE_EXPORT_ZBESTAND "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND m inner join ZBESTAND a on m.artnr = a.artnr and  m.filialnr = a.filialnr"
    sSQL = sSQL & " set m.BESTAND = a.Bestand, m.minbest = a.minbest "
    
    If gcFilNr = "1" Then
        sSQL = sSQL & " where a.filialnr = 2 "
    ElseIf gcFilNr = "2" Then
        sSQL = sSQL & " where a.filialnr = 1 "
    End If
    

    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set BESTAND =0 where bestand is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set minbest =0 where minbest is null "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    
     
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE1 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE2 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE3 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE1 = '1=' "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE1 = m.BESTANDSANGABE1 + 'B' + cStr(a.Bestand) + 'M' + cStr(a.MINBEST)"
    sSQL = sSQL & " where a.filialnr = 1 "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE2 = '2=' "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE2 = m.BESTANDSANGABE2 + 'B' + cStr(a.Bestand) + 'M' + cStr(a.MINBEST)"
    sSQL = sSQL & " where a.filialnr = 2 "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE1 = left(BESTANDSANGABE1,9) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE2 = left(BESTANDSANGABE2,9) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR1 = round(LEKPR1,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR2 = round(LEKPR2,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR3 = round(LEKPR3,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Select "
    sSQL = sSQL & " SCANCODE  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", KUERZEL1 "
    sSQL = sSQL & ", MINMEN1 "
    sSQL = sSQL & ", LEKPR1 "
    
    sSQL = sSQL & ", KUERZEL2 "
    sSQL = sSQL & ", MINMEN2 "
    sSQL = sSQL & ", LEKPR2 "
    
    sSQL = sSQL & ", KUERZEL3 "
    sSQL = sSQL & ", MINMEN3 "
    sSQL = sSQL & ", LEKPR3 "
    
    sSQL = sSQL & ", BESTANDSANGABE1 "
    sSQL = sSQL & ", BESTANDSANGABE2 "
    sSQL = sSQL & ", BESTANDSANGABE3 "

    sSQL = sSQL & " from MDE_EXPORT_SCANPAL3 order by Artnr "
    
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        sAusgabedatname = "ARTIKEL_for_MDE.txt"
        
        cdatei = cPfad1 & sAusgabedatname
        cPfad = cPfad1
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            cFeld = ""
            
            If Not IsNull(rsrs!SCANCODE) Then
                cFeld = Left(rsrs!SCANCODE, 13)
            End If
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(13 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = Left(rsrs!BEZEICH, 35)
            End If
            
            cFeld = SwapStr(cFeld, ",", ".")
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(35 - Len(cFeld)) & ","
            
            
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL1) Then
                cFeld = Left(rsrs!KUERZEL1, 2)
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR1) Then
                cFeld = Format(rsrs!LEKPR1, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If
            
            
            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN1) Then
                If Len(rsrs!MINMEN1) <= 2 Then
                    cFeld = rsrs!MINMEN1
                Else
                    cFeld = "~~"
                End If
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            
            
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE1) Then
                cFeld = rsrs!BESTANDSANGABE1
            End If

            cSatz = cSatz & Space(1) & cFeld & Space(9 - Len(cFeld)) & ","
            
            
            
            
            
            
            
            
            
            
            
            
            
            '2. Zeile
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL2) Then
                cFeld = Left(rsrs!KUERZEL2, 2)
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR2) Then
                cFeld = Format(rsrs!LEKPR2, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If
            
            
            
            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN2) Then
                If Len(rsrs!MINMEN2) <= 2 Then
                    cFeld = rsrs!MINMEN2
                Else
                    cFeld = "~~"
                End If
            End If
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE2) Then
                cFeld = rsrs!BESTANDSANGABE2
            End If

            cSatz = cSatz & Space(1) & cFeld & Space(9 - Len(cFeld)) & ","
            
            '3.

            cFeld = ""
            If Not IsNull(rsrs!KUERZEL3) Then
                cFeld = Left(rsrs!KUERZEL3, 2)
            End If

            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR3) Then
                cFeld = Format(rsrs!LEKPR3, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If

            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            

            cFeld = ""
            If Not IsNull(rsrs!MINMEN3) Then
                If Len(rsrs!MINMEN3) <= 2 Then
                    cFeld = rsrs!MINMEN3
                Else
                    cFeld = "~~"
                End If
            End If
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            
            
            
            
'            cFeld = ""
'            If Not IsNull(rsrs!BESTANDSANGABE) Then
'                cFeld = rsrs!BESTANDSANGABE
'            End If
'
'            cSatz = cSatz & cFeld
            
            
            
            
            
            
            
            
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("MDE_EXPORT_SCANPAL3", gdApp) Then
    
        Dim lTimeOut As Long
        
        lTimeOut = 200
        
        Dim iErr_Zaehler As Integer
        Dim iTimeout As Integer
        
        Dim bTimeout_erreicht As Boolean
        Dim ctmp As String
        
        bTimeout_erreicht = False
        
        iTimeout = lTimeOut * 10
        
        iErr_Zaehler = 0
        Do While FileExists(cdatei) = True
        
            anzeige "rot2", "MDE-Gerät wird betankt, bitte warten(" & iErr_Zaehler & ")...", Label1(4)
            PauseSi 0.1
            iErr_Zaehler = iErr_Zaehler + 1
            If iErr_Zaehler > iTimeout Then
                bTimeout_erreicht = True
                Exit Do
            End If
            
            
        Loop
        
        If bTimeout_erreicht = True Then
'            ctmp = "Fehler - Zeitlimit erreicht" & vbCrLf & vbCrLf

            ctmp = "Das hat leider nicht geklappt." & vbCrLf & vbCrLf
            ctmp = ctmp & "'SD Card is ready' - auf dem Display?" & vbCrLf
            ctmp = ctmp & "MDE Gerät richtig eingesteckt?" & vbCrLf & vbCrLf
            ctmp = ctmp & "Diesen Programmteil bitte über 'Schließen' verlassen und neubetreten!"
            
            iRet = MsgBox(ctmp, vbCritical + vbOKOnly, "Winkiss Hinweis")
            
            Exit Sub
        Else
        
            Pause 2
            anzeige "ERFOLG", "Das MDE-Gerät ist jetzt erfolgreich betankt.", Label1(4)
            MsgBox "Das MDE-Gerät ist jetzt erfolgreich betankt.", vbInformation, "Winkiss Hinweis:"
        
'            MsgBox "Die Dateiausgabe war erfolgreich! Die Datei: 'ARTIKEL_for_MDE.txt' befindet sich unter: " & cPfad1, vbInformation, "Winkiss Information:"
            
        End If
        
        

        
        
        
        
'        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportCSV_ScanPal_mitBestand"
        Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub ExportCSV_ScanPal_mitBestand_OnlyFil()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer
    Dim cPreis          As String
    Dim cFeld           As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    If gsTankPfad <> "" Then
        cPfad1 = gsTankPfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
    Else
        cPfad1 = gcDBPfad      'dbpfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
        cPfad1 = cPfad1 & "Box\"
    End If
    
    
    loeschNEW "MDE_EXPORT_ZBESTAND", gdApp
    
    sSQL = "Select artnr,bestand,minbest, Lastdate,bestand as vkmenge  into MDE_EXPORT_ZBESTAND from Artikel  where artnr in (Select Artnr from MDE_EXPORT_SCANPAL3)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set BESTAND =0 where bestand is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set minbest =0 where minbest is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set Lastdate =  null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set vkmenge =0 where vkmenge is null "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    loeschNEW "MDE_EXPORT_KASS", gdApp
    
    sSQL = "Select artnr,max(adate) as lastdate into MDE_EXPORT_KASS from Kassjour  " 'where artnr in (Select Artnr from MDE_EXPORT_SCANPAL3)"
    sSQL = sSQL & " group by artnr "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND m inner join MDE_EXPORT_KASS a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.lastdate = a.lastdate "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    loeschNEW "MDE_EXPORT_KASSMENGE", gdApp
    
    sSQL = "Select artnr,sum(MENGE) as vkmenge into MDE_EXPORT_KASSMENGE from Kassjour  where " 'artnr in (Select Artnr from MDE_EXPORT_SCANPAL3)"
    sSQL = sSQL & " Kassjour.adate >= " & CLng(DateValue(Now) - 30)
    sSQL = sSQL & " group by artnr "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND m inner join MDE_EXPORT_KASSMENGE a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.vkmenge = a.vkmenge "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set vkmenge =0 where vkmenge is null "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    loeschNEW "MDE_EXPORT_KASSMENGE", gdApp
   
    loeschNEW "MDE_EXPORT_KASS", gdApp
    
    
    
    
     
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE1 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE2 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE3 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE1 = ' ' "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE1 = m.BESTANDSANGABE1 + 'B' + cStr(a.Bestand) + 'M' + cStr(a.MINBEST)"
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE2 = cStr(a.lastdate) "
    sSQL = sSQL & " where not a.lastdate is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE3 = 'VKM ' "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE3 = m.BESTANDSANGABE3 + cStr(a.VKMENGE) "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE1 = left(BESTANDSANGABE1,9) "
    gdApp.Execute sSQL, dbFailOnError
    
'    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE2 = left(BESTANDSANGABE2,9) "
'    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE3 = left(BESTANDSANGABE3,9) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR1 = round(LEKPR1,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR2 = round(LEKPR2,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR3 = round(LEKPR3,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Select "
    sSQL = sSQL & " SCANCODE  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", KUERZEL1 "
    sSQL = sSQL & ", MINMEN1 "
    sSQL = sSQL & ", LEKPR1 "
    
    sSQL = sSQL & ", KUERZEL2 "
    sSQL = sSQL & ", MINMEN2 "
    sSQL = sSQL & ", LEKPR2 "
    
    sSQL = sSQL & ", KUERZEL3 "
    sSQL = sSQL & ", MINMEN3 "
    sSQL = sSQL & ", LEKPR3 "
    
    sSQL = sSQL & ", BESTANDSANGABE1 "
    sSQL = sSQL & ", BESTANDSANGABE2 "
    sSQL = sSQL & ", BESTANDSANGABE3 "

    sSQL = sSQL & " from MDE_EXPORT_SCANPAL3 order by Artnr "
    
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        sAusgabedatname = "ARTIKEL_for_MDE.txt"
        
        cdatei = cPfad1 & sAusgabedatname
        cPfad = cPfad1
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            cFeld = ""
            
            If Not IsNull(rsrs!SCANCODE) Then
                cFeld = Left(rsrs!SCANCODE, 13)
            End If
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(13 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = Left(rsrs!BEZEICH, 35)
            End If
            
            cFeld = SwapStr(cFeld, ",", ".")
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(35 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL1) Then
                cFeld = Left(rsrs!KUERZEL1, 2)
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR1) Then
                cFeld = Format(rsrs!LEKPR1, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If
            
            
            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN1) Then
                If Len(rsrs!MINMEN1) <= 2 Then
                    cFeld = rsrs!MINMEN1
                Else
                    cFeld = "~~"
                End If
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE1) Then
                cFeld = rsrs!BESTANDSANGABE1
            End If

            cSatz = cSatz & Space(1) & cFeld & Space(9 - Len(cFeld)) & ","
            
            
            
            
            '2. Zeile
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL2) Then
                cFeld = Left(rsrs!KUERZEL2, 2)
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR2) Then
                cFeld = Format(rsrs!LEKPR2, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If
            
            
            
            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN2) Then
                If Len(rsrs!MINMEN2) <= 2 Then
                    cFeld = rsrs!MINMEN2
                Else
                    cFeld = "~~"
                End If
            End If
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE2) Then
                cFeld = Trim(rsrs!BESTANDSANGABE2)
                
                cFeld = Format$(cFeld, "DD.MM.YY")
                
                
            End If

            cSatz = cSatz & Space(2) & cFeld & Space(8 - Len(cFeld)) & ","
            
            '3.

            cFeld = ""
            If Not IsNull(rsrs!KUERZEL3) Then
                cFeld = Left(rsrs!KUERZEL3, 2)
            End If

            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR3) Then
                cFeld = Format(rsrs!LEKPR3, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If

            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            

            cFeld = ""
            If Not IsNull(rsrs!MINMEN3) Then
                If Len(rsrs!MINMEN3) <= 2 Then
                    cFeld = rsrs!MINMEN3
                Else
                    cFeld = "~~"
                End If
            End If
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE3) Then
                cFeld = rsrs!BESTANDSANGABE3
            End If

            cSatz = cSatz & Space(1) & cFeld & Space(9 - Len(cFeld))
            
          
            
            
            
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("MDE_EXPORT_SCANPAL3", gdApp) Then
    
        Dim lTimeOut As Long
        
        lTimeOut = 200
        
        Dim iErr_Zaehler As Integer
        Dim iTimeout As Integer
        
        Dim bTimeout_erreicht As Boolean
        Dim ctmp As String
        
        bTimeout_erreicht = False
        
        iTimeout = lTimeOut * 10
        
        iErr_Zaehler = 0
        Do While FileExists(cdatei) = True
        
            anzeige "rot2", "MDE-Gerät wird betankt, bitte warten(" & iErr_Zaehler & ")...", Label1(4)
            PauseSi 0.1
            iErr_Zaehler = iErr_Zaehler + 1
            If iErr_Zaehler > iTimeout Then
                bTimeout_erreicht = True
                Exit Do
            End If
            
            
        Loop
        
        If bTimeout_erreicht = True Then
'            ctmp = "Fehler - Zeitlimit erreicht" & vbCrLf & vbCrLf

            ctmp = "Das hat leider nicht geklappt." & vbCrLf & vbCrLf
            ctmp = ctmp & "'SD Card is ready' - auf dem Display?" & vbCrLf
            ctmp = ctmp & "MDE Gerät richtig eingesteckt?" & vbCrLf & vbCrLf
            ctmp = ctmp & "Diesen Programmteil bitte über 'Schließen' verlassen und neubetreten!"
            
            iRet = MsgBox(ctmp, vbCritical + vbOKOnly, "Winkiss Hinweis")
            
            Exit Sub
        Else
        
            Pause 2
            anzeige "ERFOLG", "Das MDE-Gerät ist jetzt erfolgreich betankt.", Label1(4)
            MsgBox "Das MDE-Gerät ist jetzt erfolgreich betankt.", vbInformation, "Winkiss Hinweis:"
        
'            MsgBox "Die Dateiausgabe war erfolgreich! Die Datei: 'ARTIKEL_for_MDE.txt' befindet sich unter: " & cPfad1, vbInformation, "Winkiss Information:"
            
        End If
        
        

        
        
        
        
'        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportCSV_ScanPal_mitBestand_OnlyFil"
        Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub ExportCSV_ScanPal_mitBestandKVK_OnlyFil()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer
    Dim cPreis          As String
    Dim cFeld           As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    If gsTankPfad <> "" Then
        cPfad1 = gsTankPfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
    Else
        cPfad1 = gcDBPfad      'dbpfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
        cPfad1 = cPfad1 & "Box\"
    End If
    
    
    loeschNEW "MDE_EXPORT_ZBESTAND", gdApp
    
    sSQL = "Select artnr,bestand,KVKPR1, Lastdate,bestand as vkmenge  into MDE_EXPORT_ZBESTAND from Artikel  where artnr in (Select Artnr from MDE_EXPORT_SCANPAL3)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set BESTAND =0 where bestand is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set KVKPR1 =0 where KVKPR1 is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set KVKPR1 = round(KVKPR1,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set Lastdate =  null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set vkmenge =0 where vkmenge is null "
    gdApp.Execute sSQL, dbFailOnError
    
    
    loeschNEW "MDE_EXPORT_KASS", gdApp
    
    sSQL = "Select artnr,max(adate) as lastdate into MDE_EXPORT_KASS from Kassjour  " 'where artnr in (Select Artnr from MDE_EXPORT_SCANPAL3)"
    sSQL = sSQL & " group by artnr "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND m inner join MDE_EXPORT_KASS a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.lastdate = a.lastdate "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    loeschNEW "MDE_EXPORT_KASSMENGE", gdApp
    
    sSQL = "Select artnr,sum(MENGE) as vkmenge into MDE_EXPORT_KASSMENGE from Kassjour  where " 'artnr in (Select Artnr from MDE_EXPORT_SCANPAL3)"
    sSQL = sSQL & " Kassjour.adate >= " & CLng(DateValue(Now) - 30)
    sSQL = sSQL & " group by artnr "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND m inner join MDE_EXPORT_KASSMENGE a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.vkmenge = a.vkmenge "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set vkmenge =0 where vkmenge is null "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    loeschNEW "MDE_EXPORT_KASSMENGE", gdApp
   
    loeschNEW "MDE_EXPORT_KASS", gdApp
    
    
    
    
     
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE1 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE2 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE3 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE1 = ' ' "
    gdApp.Execute sSQL, dbFailOnError
    
    
'    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
'    sSQL = sSQL & " set m.BESTANDSANGABE1 = m.BESTANDSANGABE1 + 'B' + cStr(a.Bestand) + ' ' + cStr(a.KVKPR1)"
'    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE1 = m.BESTANDSANGABE1 + 'B' + cStr(a.Bestand) + ' ' + cStr(Format(a.KVKPR1, '##0.00'))"
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE2 = cStr(a.lastdate) "
    sSQL = sSQL & " where not a.lastdate is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE3 = 'VKM ' "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE3 = m.BESTANDSANGABE3 + cStr(a.VKMENGE) "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE1 = left(BESTANDSANGABE1,9) "
    gdApp.Execute sSQL, dbFailOnError
    
'    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE2 = left(BESTANDSANGABE2,9) "
'    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE3 = left(BESTANDSANGABE3,9) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR1 = round(LEKPR1,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR2 = round(LEKPR2,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR3 = round(LEKPR3,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Select "
    sSQL = sSQL & " SCANCODE  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", KUERZEL1 "
    sSQL = sSQL & ", MINMEN1 "
    sSQL = sSQL & ", LEKPR1 "
    
    sSQL = sSQL & ", KUERZEL2 "
    sSQL = sSQL & ", MINMEN2 "
    sSQL = sSQL & ", LEKPR2 "
    
    sSQL = sSQL & ", KUERZEL3 "
    sSQL = sSQL & ", MINMEN3 "
    sSQL = sSQL & ", LEKPR3 "
    
    sSQL = sSQL & ", BESTANDSANGABE1 "
    sSQL = sSQL & ", BESTANDSANGABE2 "
    sSQL = sSQL & ", BESTANDSANGABE3 "

    sSQL = sSQL & " from MDE_EXPORT_SCANPAL3 order by Artnr "
    
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        sAusgabedatname = "ARTIKEL_for_MDE.txt"
        
        cdatei = cPfad1 & sAusgabedatname
        cPfad = cPfad1
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            cFeld = ""
            
            If Not IsNull(rsrs!SCANCODE) Then
                cFeld = Left(rsrs!SCANCODE, 13)
            End If
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(13 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = Left(rsrs!BEZEICH, 35)
            End If
            
            cFeld = SwapStr(cFeld, ",", ".")
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(35 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL1) Then
                cFeld = Left(rsrs!KUERZEL1, 2)
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR1) Then
                cFeld = Format(rsrs!LEKPR1, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If
            
            
            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN1) Then
                If Len(rsrs!MINMEN1) <= 2 Then
                    cFeld = rsrs!MINMEN1
                Else
                    cFeld = "~~"
                End If
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE1) Then
                cFeld = rsrs!BESTANDSANGABE1
                cFeld = Format(cFeld, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
            End If

            cSatz = cSatz & Space(1) & cFeld & Space(9 - Len(cFeld)) & ","
            
            
            
            
            '2. Zeile
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL2) Then
                cFeld = Left(rsrs!KUERZEL2, 2)
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR2) Then
                cFeld = Format(rsrs!LEKPR2, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If
            
            
            
            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN2) Then
                If Len(rsrs!MINMEN2) <= 2 Then
                    cFeld = rsrs!MINMEN2
                Else
                    cFeld = "~~"
                End If
            End If
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE2) Then
                cFeld = Trim(rsrs!BESTANDSANGABE2)
                
                cFeld = Format$(cFeld, "DD.MM.YY")
                
                
            End If

            cSatz = cSatz & Space(2) & cFeld & Space(8 - Len(cFeld)) & ","
            
            '3.

            cFeld = ""
            If Not IsNull(rsrs!KUERZEL3) Then
                cFeld = Left(rsrs!KUERZEL3, 2)
            End If

            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR3) Then
                cFeld = Format(rsrs!LEKPR3, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If

            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            

            cFeld = ""
            If Not IsNull(rsrs!MINMEN3) Then
                If Len(rsrs!MINMEN3) <= 2 Then
                    cFeld = rsrs!MINMEN3
                Else
                    cFeld = "~~"
                End If
            End If
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE3) Then
                cFeld = rsrs!BESTANDSANGABE3
            End If

            cSatz = cSatz & Space(1) & cFeld & Space(9 - Len(cFeld))
            
          
            
            
            
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("MDE_EXPORT_SCANPAL3", gdApp) Then
    
        Dim lTimeOut As Long
        lTimeOut = 200
        
        Dim iErr_Zaehler As Integer
        Dim iTimeout As Integer
        
        Dim bTimeout_erreicht As Boolean
        Dim ctmp As String
        
        bTimeout_erreicht = False
        iTimeout = lTimeOut * 10
        
        iErr_Zaehler = 0
        Do While FileExists(cdatei) = True
            anzeige "rot2", "MDE-Gerät wird betankt, bitte warten(" & iErr_Zaehler & ")...", Label1(4)
            PauseSi 0.1
            iErr_Zaehler = iErr_Zaehler + 1
            If iErr_Zaehler > iTimeout Then
                bTimeout_erreicht = True
                Exit Do
            End If
        Loop
        
        If bTimeout_erreicht = True Then
            ctmp = "Das hat leider nicht geklappt." & vbCrLf & vbCrLf
            ctmp = ctmp & "'SD Card is ready' - auf dem Display?" & vbCrLf
            ctmp = ctmp & "MDE Gerät richtig eingesteckt?" & vbCrLf & vbCrLf
            ctmp = ctmp & "Diesen Programmteil bitte über 'Schließen' verlassen und neubetreten!"
            
            iRet = MsgBox(ctmp, vbCritical + vbOKOnly, "Winkiss Hinweis")
            Exit Sub
        Else
            Pause 2
            anzeige "ERFOLG", "Das MDE-Gerät ist jetzt erfolgreich betankt.", Label1(4)
            MsgBox "Das MDE-Gerät ist jetzt erfolgreich betankt.", vbInformation, "Winkiss Hinweis:"
        End If
    Else
        anzeige "rot", "Keine Daten vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportCSV_ScanPal_mitBestandKVK_OnlyFil"
        Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub ExportCSV_ScanPal_mitKVK_OnlyFil()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer
    Dim cPreis          As String
    Dim cFeld           As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    If gsTankPfad <> "" Then
        cPfad1 = gsTankPfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
    Else
        cPfad1 = gcDBPfad      'dbpfad
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
        cPfad1 = cPfad1 & "Box\"
    End If
    
    
    loeschNEW "MDE_EXPORT_ZBESTAND", gdApp
    
    sSQL = "Select artnr,bestand,minbest,KVKPR1, Lastdate,bestand as zumenge  into MDE_EXPORT_ZBESTAND from Artikel  where artnr in (Select Artnr from MDE_EXPORT_SCANPAL3)"
    gdApp.Execute sSQL, dbFailOnError
    

    sSQL = "Update MDE_EXPORT_ZBESTAND set KVKPR1 =0 where KVKPR1 is null "
    gdApp.Execute sSQL, dbFailOnError

    sSQL = "Update MDE_EXPORT_ZBESTAND set Lastdate =  null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set zumenge =0 where zumenge is null "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    loeschNEW "MDE_EXPORT_ZU", gdApp
    
    
    
    
    
    

    
    
    
    
    
    
    
    
    
    sSQL = "Select artnr,max(adate) as lastdate into MDE_EXPORT_ZU from ZUGANG  "
    sSQL = sSQL & " group by artnr "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND m inner join MDE_EXPORT_ZU a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.lastdate = a.lastdate "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    loeschNEW "MDE_EXPORT_ZUMENGE", gdApp
    
    sSQL = "Select z.artnr,sum(BEWEGUNG) as zumenge into MDE_EXPORT_ZUMENGE from ZUGANG z inner join MDE_EXPORT_ZBESTAND m on z.adate = m.lastdate "
    sSQL = sSQL & " and z.artnr = m.artnr "
    sSQL = sSQL & " group by z.artnr "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND m inner join MDE_EXPORT_ZUMENGE a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.zumenge = a.zumenge "
    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set zumenge =0 where zumenge is null "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    loeschNEW "MDE_EXPORT_ZUMENGE", gdApp
   
    loeschNEW "MDE_EXPORT_ZU", gdApp
    
    
    
    
     
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE1 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE2 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table MDE_EXPORT_SCANPAL3 add  BESTANDSANGABE3 Text(40)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE1 = ' ' "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_ZBESTAND set kvkpr1 = round(kvkpr1,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE1 = m.BESTANDSANGABE1 + cStr(a.kvkpr1) "
    gdApp.Execute sSQL, dbFailOnError
    
'    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
'    sSQL = sSQL & " set m.BESTANDSANGABE1 = m.BESTANDSANGABE1 + 'B' + cStr(a.Bestand) + 'M' + cStr(a.MINBEST)"
'    gdApp.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE2 = cStr(a.lastdate) "
    sSQL = sSQL & " where not a.lastdate is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE3 = 'ZuM ' "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 m inner join MDE_EXPORT_ZBESTAND a on m.artnr = a.artnr "
    sSQL = sSQL & " set m.BESTANDSANGABE3 = m.BESTANDSANGABE3 + cStr(a.ZUMENGE) "
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE1 = left(BESTANDSANGABE1,9) "
    gdApp.Execute sSQL, dbFailOnError
    
'    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE2 = left(BESTANDSANGABE2,9) "
'    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set BESTANDSANGABE3 = left(BESTANDSANGABE3,9) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR1 = round(LEKPR1,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR2 = round(LEKPR2,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update MDE_EXPORT_SCANPAL3 set LEKPR3 = round(LEKPR3,2) "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Select "
    sSQL = sSQL & " SCANCODE  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", KUERZEL1 "
    sSQL = sSQL & ", MINMEN1 "
    sSQL = sSQL & ", LEKPR1 "
    
    sSQL = sSQL & ", KUERZEL2 "
    sSQL = sSQL & ", MINMEN2 "
    sSQL = sSQL & ", LEKPR2 "
    
    sSQL = sSQL & ", KUERZEL3 "
    sSQL = sSQL & ", MINMEN3 "
    sSQL = sSQL & ", LEKPR3 "
    
    sSQL = sSQL & ", BESTANDSANGABE1 "
    sSQL = sSQL & ", BESTANDSANGABE2 "
    sSQL = sSQL & ", BESTANDSANGABE3 "

    sSQL = sSQL & " from MDE_EXPORT_SCANPAL3 order by Artnr "
    
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        sAusgabedatname = "ARTIKEL_for_MDE.txt"
        
        cdatei = cPfad1 & sAusgabedatname
        cPfad = cPfad1
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            cFeld = ""
            
            If Not IsNull(rsrs!SCANCODE) Then
                cFeld = Left(rsrs!SCANCODE, 13)
            End If
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(13 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = Left(rsrs!BEZEICH, 35)
            End If
            
            cFeld = SwapStr(cFeld, ",", ".")
            
            cSatz = cSatz & cFeld
            cSatz = cSatz & Space(35 - Len(cFeld)) & ","
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL1) Then
                cFeld = Left(rsrs!KUERZEL1, 2)
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR1) Then
                cFeld = Format(rsrs!LEKPR1, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If
            
            
            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN1) Then
                If Len(rsrs!MINMEN1) <= 2 Then
                    cFeld = rsrs!MINMEN1
                Else
                    cFeld = "~~"
                End If
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE1) Then
                cFeld = rsrs!BESTANDSANGABE1
                cFeld = Format(cFeld, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
            End If

            cSatz = cSatz & Space(1) & cFeld & Space(9 - Len(cFeld)) & ","
            
            
            
            
            '2. Zeile
            
            cFeld = ""
            If Not IsNull(rsrs!KUERZEL2) Then
                cFeld = Left(rsrs!KUERZEL2, 2)
            End If
            
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR2) Then
                cFeld = Format(rsrs!LEKPR2, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If
            
            
            
            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            
            cFeld = ""
            If Not IsNull(rsrs!MINMEN2) Then
                If Len(rsrs!MINMEN2) <= 2 Then
                    cFeld = rsrs!MINMEN2
                Else
                    cFeld = "~~"
                End If
            End If
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE2) Then
                cFeld = Trim(rsrs!BESTANDSANGABE2)
                
                cFeld = Format$(cFeld, "DD.MM.YY")
                
                
            End If

            cSatz = cSatz & Space(2) & cFeld & Space(8 - Len(cFeld)) & ","
            
            '3.

            cFeld = ""
            If Not IsNull(rsrs!KUERZEL3) Then
                cFeld = Left(rsrs!KUERZEL3, 2)
            End If

            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            cFeld = ""
            If Not IsNull(rsrs!LEKPR3) Then
                cFeld = Format(rsrs!LEKPR3, "##0.00")
                cFeld = SwapStr(cFeld, ",", ".")
                
                If Len(cFeld) > 5 Then
                    cFeld = "~~~~~"
                End If
            End If

            cSatz = cSatz & Space(5 - Len(cFeld)) & cFeld & Space(1)
            

            cFeld = ""
            If Not IsNull(rsrs!MINMEN3) Then
                If Len(rsrs!MINMEN3) <= 2 Then
                    cFeld = rsrs!MINMEN3
                Else
                    cFeld = "~~"
                End If
            End If
            cSatz = cSatz & Space(2 - Len(cFeld)) & cFeld
            
            
            cFeld = ""
            If Not IsNull(rsrs!BESTANDSANGABE3) Then
                cFeld = rsrs!BESTANDSANGABE3
            End If

            cSatz = cSatz & Space(1) & cFeld & Space(9 - Len(cFeld))
            
          
            
            
            
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("MDE_EXPORT_SCANPAL3", gdApp) Then
    
        Dim lTimeOut As Long
        
        lTimeOut = 200
        
        Dim iErr_Zaehler As Integer
        Dim iTimeout As Integer
        
        Dim bTimeout_erreicht As Boolean
        Dim ctmp As String
        
        bTimeout_erreicht = False
        
        iTimeout = lTimeOut * 10
        
        iErr_Zaehler = 0
        Do While FileExists(cdatei) = True
        
            anzeige "rot2", "MDE-Gerät wird betankt, bitte warten(" & iErr_Zaehler & ")...", Label1(4)
            PauseSi 0.1
            iErr_Zaehler = iErr_Zaehler + 1
            If iErr_Zaehler > iTimeout Then
                bTimeout_erreicht = True
                Exit Do
            End If
            
            
        Loop
        
        If bTimeout_erreicht = True Then
'            ctmp = "Fehler - Zeitlimit erreicht" & vbCrLf & vbCrLf

            ctmp = "Das hat leider nicht geklappt." & vbCrLf & vbCrLf
            ctmp = ctmp & "'SD Card is ready' - auf dem Display?" & vbCrLf
            ctmp = ctmp & "MDE Gerät richtig eingesteckt?" & vbCrLf & vbCrLf
            ctmp = ctmp & "Diesen Programmteil bitte über 'Schließen' verlassen und neubetreten!"
            
            iRet = MsgBox(ctmp, vbCritical + vbOKOnly, "Winkiss Hinweis")
            
            Exit Sub
        Else
        
            Pause 2
            anzeige "ERFOLG", "Das MDE-Gerät ist jetzt erfolgreich betankt.", Label1(4)
            MsgBox "Das MDE-Gerät ist jetzt erfolgreich betankt.", vbInformation, "Winkiss Hinweis:"
        
'            MsgBox "Die Dateiausgabe war erfolgreich! Die Datei: 'ARTIKEL_for_MDE.txt' befindet sich unter: " & cPfad1, vbInformation, "Winkiss Information:"
            
        End If
        
        

        
        
        
        
'        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportCSV_ScanPal_mitKVK_OnlyFil"
        Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "MAXVPE", gdBase
    loeschNEW "MDE_EXPORT_REWE", gdBase

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

Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If Index = 0 Then
        LiefKuerzelAufloesung lbl6(1), Text1(0)
    End If
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim sAuswahlfeld As String
    Dim ctmp As String
    
    
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            
            Case Is = 0
                gF2Prompt.cFeld = "LINR"
                
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
'                    Label1(10).Caption = gF2Prompt.cWert
                End If
            
        End Select
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil MDE betanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
