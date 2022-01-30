VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL215 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Datenschutzblatt konfigurieren"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check4 
      Caption         =   "Geburtstag drucken"
      Height          =   375
      Left            =   9720
      TabIndex        =   41
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Meldung, wenn DS unterschrieben"
      Height          =   615
      Left            =   9720
      TabIndex        =   40
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Mobil"
      Height          =   255
      Index           =   8
      Left            =   9720
      TabIndex        =   39
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Tel"
      Height          =   255
      Index           =   7
      Left            =   9720
      TabIndex        =   38
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Mail"
      Height          =   255
      Index           =   6
      Left            =   9720
      TabIndex        =   37
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Geburtsdatum"
      Height          =   255
      Index           =   5
      Left            =   9720
      TabIndex        =   36
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Strasse"
      Height          =   255
      Index           =   4
      Left            =   9720
      TabIndex        =   35
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "PLZ"
      Height          =   255
      Index           =   3
      Left            =   9720
      TabIndex        =   34
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Ort"
      Height          =   255
      Index           =   2
      Left            =   9720
      TabIndex        =   33
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Vorname"
      Height          =   255
      Index           =   1
      Left            =   9720
      TabIndex        =   32
      Top             =   840
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   9720
      TabIndex        =   30
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   29
      Top             =   6720
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   28
      Top             =   7080
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   27
      Top             =   5760
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   26
      Top             =   6120
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   12
      Left            =   1200
      MaxLength       =   250
      TabIndex        =   25
      Top             =   4560
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   720
      TabIndex        =   24
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Schriftgröße 9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   9720
      TabIndex        =   23
      Top             =   3240
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Schriftgröße 10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9720
      TabIndex        =   22
      Top             =   2880
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "automatisch bei Kundenauswahl an der Kasse drucken (wenn DS noch nicht unterschrieben)"
      Height          =   1215
      Left            =   9720
      TabIndex        =   21
      Top             =   3960
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   720
      TabIndex        =   6
      Top             =   4200
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtElement 
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
      Index           =   19
      Left            =   120
      TabIndex        =   20
      Top             =   7920
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   19
      Top             =   5160
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   11
      Left            =   1200
      MaxLength       =   250
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   1200
      MaxLength       =   250
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   1200
      MaxLength       =   250
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   14
      Text            =   "frmWKL215.frx":0000
      Top             =   2880
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2520
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
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
      Index           =   8
      Left            =   120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3480
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   11
      Text            =   "frmWKL215.frx":0006
      Top             =   1920
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1560
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   9
      Text            =   "frmWKL215.frx":000C
      Top             =   960
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   600
      Width           =   9255
   End
   Begin VB.TextBox txtElement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   120
      Width           =   9255
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   3720
      Width           =   495
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   3
      Left            =   9720
      TabIndex        =   0
      Top             =   7800
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   0
      Left            =   9720
      TabIndex        =   3
      Top             =   7200
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   1
      Left            =   9720
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Standard"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   2
      Left            =   9720
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Testdruck"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pflichtfeldkennzeichnung mit Sternchen (*)"
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
      Index           =   4
      Left            =   9720
      TabIndex        =   31
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmWKL215"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    Select Case Index


        Case 0 'speichern
            speicher_Datenschutz
            zeige_Datenschutz
        Case 1 'Standard
            setze_standard
            speicher_Datenschutz
            zeige_Datenschutz
        
        Case 2 'Testdruck
           Dim cKdnr As String
            Set rsrs = gdBase.OpenRecordset("Select max(KUNDNR) as maxi from KUNDEN")
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                If Not IsNull(rsrs!maxi) Then
                    cKdnr = rsrs!maxi
                End If
            
            End If
            rsrs.Close: Set rsrs = Nothing
        
            DatenschutzblattKundeDrucken cKdnr
        Case 3 'schließen
            Unload frmWKL215
    End Select

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Datenschutz ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Skalieren Me, True, True: Schrift Me:
    Farbform Me, Nothing
    LogtoStart Me
    
    If NewTableSuchenDBKombi("DATENSCHUTZ", gdBase) = False Then
        setze_standard
        speicher_Datenschutz
    End If
    
    zeige_Datenschutz
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Datenschutz ist ein Fehler aufgetreten."
    
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
Private Sub zeige_Datenschutz()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from Datenschutz"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Element1) Then
            txtElement(1).Text = rsrs!Element1
        End If
        If Not IsNull(rsrs!Element2) Then
            txtElement(2).Text = rsrs!Element2
        End If
        If Not IsNull(rsrs!Element3) Then
            txtElement(3).Text = rsrs!Element3
        End If
        If Not IsNull(rsrs!Element4) Then
            txtElement(4).Text = rsrs!Element4
        End If
        If Not IsNull(rsrs!Element5) Then
            txtElement(5).Text = rsrs!Element5
        End If
        If Not IsNull(rsrs!Element6) Then
            txtElement(6).Text = rsrs!Element6
        End If
        If Not IsNull(rsrs!Element7) Then
            txtElement(7).Text = rsrs!Element7
        End If
        If Not IsNull(rsrs!Element8) Then
            txtElement(8).Text = rsrs!Element8
        End If
        If Not IsNull(rsrs!Element9) Then
            txtElement(9).Text = rsrs!Element9
        End If
        If Not IsNull(rsrs!Element10) Then
            txtElement(10).Text = rsrs!Element10
        End If
        If Not IsNull(rsrs!Element11) Then
            txtElement(11).Text = rsrs!Element11
        End If
        If Not IsNull(rsrs!Element12) Then
            txtElement(12).Text = rsrs!Element12
        End If
        If Not IsNull(rsrs!Element13) Then
            txtElement(13).Text = rsrs!Element13
        End If
        If Not IsNull(rsrs!Element14) Then
            txtElement(14).Text = rsrs!Element14
        End If
        
        If Not IsNull(rsrs!Element15) Then
            txtElement(15).Text = rsrs!Element15
        End If
        
        If Not IsNull(rsrs!Element16) Then
            txtElement(16).Text = rsrs!Element16
        End If
        
        If Not IsNull(rsrs!Element17) Then
            txtElement(17).Text = rsrs!Element17
        End If
        
        If Not IsNull(rsrs!Element18) Then
            txtElement(18).Text = rsrs!Element18
        End If
        
        If Not IsNull(rsrs!Element19) Then
            txtElement(19).Text = rsrs!Element19
        End If
        
        If Not IsNull(rsrs!PflichtName) Then
            If rsrs!PflichtName Then
                Check2(0).Value = vbChecked
            Else
                Check2(0).Value = vbUnchecked
            End If
        End If
        
        If Not IsNull(rsrs!PflichtvorName) Then
            If rsrs!PflichtvorName Then
                Check2(1).Value = vbChecked
            Else
                Check2(1).Value = vbUnchecked
            End If
        End If
        
        If Not IsNull(rsrs!Pflichtstadt) Then
            If rsrs!Pflichtstadt Then
                Check2(2).Value = vbChecked
            Else
                Check2(2).Value = vbUnchecked
            End If
        End If
        
        
        
        
        
        If Not IsNull(rsrs!PflichtPLZ) Then
            If rsrs!PflichtPLZ Then
                Check2(3).Value = vbChecked
            Else
                Check2(3).Value = vbUnchecked
            End If
        End If
        
        If Not IsNull(rsrs!PflichtSTRASSE) Then
            If rsrs!PflichtSTRASSE Then
                Check2(4).Value = vbChecked
            Else
                Check2(4).Value = vbUnchecked
            End If
        End If
        
        If Not IsNull(rsrs!PflichtGEBDATUM) Then
            If rsrs!PflichtGEBDATUM Then
                Check2(5).Value = vbChecked
            Else
                Check2(5).Value = vbUnchecked
            End If
        End If
        
        
        
        
        If Not IsNull(rsrs!PflichtMAIL) Then
            If rsrs!PflichtMAIL Then
                Check2(6).Value = vbChecked
            Else
                Check2(6).Value = vbUnchecked
            End If
        End If
        
        If Not IsNull(rsrs!PflichtTEL) Then
            If rsrs!PflichtTEL Then
                Check2(7).Value = vbChecked
            Else
                Check2(7).Value = vbUnchecked
            End If
        End If
        
        If Not IsNull(rsrs!PflichtMOBIL) Then
            If rsrs!PflichtMOBIL Then
                Check2(8).Value = vbChecked
            Else
                Check2(8).Value = vbUnchecked
            End If
        End If
        
        
        
        
        
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If gbDSDRUCKEN Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If
    
    If gbDS_GEB_DRUCKEN Then
        Check4.Value = vbChecked
    Else
        Check4.Value = vbUnchecked
    End If
    
    If gbDSMeldungErfolg Then
        Check3.Value = vbChecked
    Else
        Check3.Value = vbUnchecked
    End If
    
    If gbDSKLEIN Then
        Option2(0).Value = True
    Else
        Option2(0).Value = False
    End If
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Datenschutz"
    Fehler.gsFehlertext = "Im Programmteil Datenschutz ist ein Fehler aufgetreten."
End Sub
Private Sub setze_standard()
    On Error GoTo LOKAL_ERROR

        txtElement(1).Text = ""
        txtElement(2).Text = ""
        txtElement(3).Text = ""
        txtElement(4).Text = ""
        txtElement(5).Text = ""
        txtElement(6).Text = ""
        txtElement(7).Text = ""
        txtElement(8).Text = ""
        txtElement(9).Text = ""
        txtElement(10).Text = ""
        txtElement(11).Text = ""
        txtElement(12).Text = ""
        txtElement(13).Text = ""
        txtElement(14).Text = ""
        
        txtElement(15).Text = ""
        txtElement(16).Text = ""
        txtElement(17).Text = ""
        txtElement(18).Text = ""
        txtElement(19).Text = ""

        Dim sFirmaName As String
        Dim sFirmaPLZ As String
        Dim sFirmaORT As String
        Dim sFirmaSTRASSE As String
        
        sFirmaName = gFirma.FirmaName
        sFirmaPLZ = gFirma.Plz
        sFirmaORT = gFirma.Ort
        sFirmaSTRASSE = gFirma.strasse


        txtElement(1).Text = "Kundenkarten-Antrag"

        txtElement(2).Text = ""
        txtElement(3).Text = "Hiermit beantrage ich eine Kundenkarte der Firma " & sFirmaName & ", gültig und einsetzbar in jeder Filiale der Firma " & sFirmaName & "." & vbCrLf
        txtElement(3).Text = txtElement(3).Text & "Mit der Teilnahme an dem Kundenkarten-Programm gehe ich keinerlei Verpflichtungen ein."

        txtElement(4).Text = "Datenschutzrechtliche Einwilligungserklärung"

        txtElement(5).Text = "Wenn Sie an unserem Kundenkarten-Programm teilnehmen, werden Ihre personenbezogenen "
        txtElement(5).Text = txtElement(5).Text & "Daten (Name, vollständige Anschrift, Geburtsdatum) Ihre "
        txtElement(5).Text = txtElement(5).Text & "Kontaktdaten (email-Adresse, Telefonnummer) sowie Ihre Einkaufsdaten "
        txtElement(5).Text = txtElement(5).Text & "von der Firma " & sFirmaName & ", vertreten durch die Firma "
        txtElement(5).Text = txtElement(5).Text & sFirmaName & "," & sFirmaSTRASSE & " in " & sFirmaPLZ & " " & sFirmaORT & " zum "
        txtElement(5).Text = txtElement(5).Text & "Zwecke der Abwicklung des Kundenkarten-Programms" & vbCrLf
        txtElement(5).Text = txtElement(5).Text & "sowie der Koordination der Termine unserer Kabinenkunden erhoben, gespeichert und genutzt." & vbCrLf
        txtElement(5).Text = txtElement(5).Text & "(nicht zutreffendes streichen)"

        txtElement(6).Text = ""

        txtElement(7).Text = "Wenn Sie darüber hinaus möchten, dass wir Sie über unsere aktuellen Aktionen, "
        txtElement(7).Text = txtElement(7).Text & "Angebote und Produkte informieren und beraten, dann teilen Sie uns doch einfach Ihr "
        txtElement(7).Text = txtElement(7).Text & "Einverständnis mit. Die Einverständniserklärung erfolgt selbstverständlich freiwillig."

        txtElement(8).Text = "Ja, ich wünsche Informationen über aktuelle Aktionen, Angebote und Produkte"
        txtElement(9).Text = "per Post"
        txtElement(10).Text = "per Email und SMS"
        txtElement(11).Text = "per Telefon"
        txtElement(12).Text = ""
        txtElement(13).Text = ""
        txtElement(14).Text = "Wenn Sie künftig unsere interessanten Informationen und Angebote nicht mehr erhalten möchten, "
        txtElement(14).Text = txtElement(14).Text & "können Sie der Verwendung Ihrer Daten für Werbezwecke jederzeit mit "
        txtElement(14).Text = txtElement(14).Text & "Wirkung für die Zukunft widersprechen. Teilen Sie uns dies bitte möglichst "
        txtElement(14).Text = txtElement(14).Text & " schriftlich an die Firma " & sFirmaName & ", vertreten durch die Firma "
        txtElement(14).Text = txtElement(14).Text & sFirmaName & "," & sFirmaSTRASSE & " in " & sFirmaPLZ & " " & sFirmaORT & " mit."

        txtElement(15).Text = ""
        txtElement(16).Text = ""
        txtElement(17).Text = ""
        txtElement(18).Text = ""
        txtElement(19).Text = ""
        
        Check2(0).Value = vbChecked
        Check2(1).Value = vbChecked
        Check2(2).Value = vbChecked
        
        Check2(3).Value = vbChecked
        Check2(4).Value = vbChecked
        Check2(5).Value = vbChecked
        
        Check2(6).Value = vbChecked
        Check2(7).Value = vbChecked
        Check2(8).Value = vbChecked

        Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "setze_standard"
    Fehler.gsFehlertext = "Im Programmteil Datenschutz ist ein Fehler aufgetreten."
End Sub
Private Sub speicher_Datenschutz()
        On Error GoTo LOKAL_ERROR

        Dim sSQL As String
        
        loeschNEW "DATENSCHUTZ", gdBase
        CreateTableT2 "DATENSCHUTZ", gdBase

        sSQL = "Insert into DATENSCHUTZ (   "
        sSQL = sSQL & " Element1 "
        sSQL = sSQL & ", Element2 "
        sSQL = sSQL & ", Element3 "
        sSQL = sSQL & ", Element4 "
        sSQL = sSQL & ", Element5 "
        sSQL = sSQL & ", Element6 "
        sSQL = sSQL & ", Element7 "
        sSQL = sSQL & ", Element8 "
        sSQL = sSQL & ", Element9 "
        sSQL = sSQL & ", Element10 "
        sSQL = sSQL & ", Element11 "
        sSQL = sSQL & ", Element12 "
        sSQL = sSQL & ", Element13 "
        sSQL = sSQL & ", Element14 "
        
        sSQL = sSQL & ", Element15 "
        sSQL = sSQL & ", Element16 "
        sSQL = sSQL & ", Element17 "
        sSQL = sSQL & ", Element18 "
        sSQL = sSQL & ", Element19 "
        
        sSQL = sSQL & ", PflichtNAME  "
        sSQL = sSQL & ", PflichtVORNAME  "
        sSQL = sSQL & ", PflichtSTADT  "
        sSQL = sSQL & ", PflichtPLZ  "
        sSQL = sSQL & ", PflichtSTRASSE  "
        sSQL = sSQL & ", PflichtGEBDATUM  "
        sSQL = sSQL & ", PflichtMAIL  "
        sSQL = sSQL & ", PflichtTEL  "
        sSQL = sSQL & ", PflichtMOBIL  "
        
        sSQL = sSQL & " ) values ( "

        sSQL = sSQL & " '" & txtElement(1).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(2).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(3).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(4).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(5).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(6).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(7).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(8).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(9).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(10).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(11).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(12).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(13).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(14).Text & "'  "
        
        sSQL = sSQL & ", '" & txtElement(15).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(16).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(17).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(18).Text & "'  "
        sSQL = sSQL & ", '" & txtElement(19).Text & "'  "
        
        
        sSQL = sSQL & ", '" & Check2(0).Value & "'  "
        sSQL = sSQL & ", '" & Check2(1).Value & "'  "
        sSQL = sSQL & ", '" & Check2(2).Value & "'  "
        
        sSQL = sSQL & ", '" & Check2(3).Value & "'  "
        sSQL = sSQL & ", '" & Check2(4).Value & "'  "
        sSQL = sSQL & ", '" & Check2(5).Value & "'  "
        
        sSQL = sSQL & ", '" & Check2(6).Value & "'  "
        sSQL = sSQL & ", '" & Check2(7).Value & "'  "
        sSQL = sSQL & ", '" & Check2(8).Value & "'  "
        
        
        
        sSQL = sSQL & ") "
        gdBase.Execute sSQL, dbFailOnError
        
        If Check1.Value = vbChecked Then
            sSQL = "Update KASSEIN Set DSDRUCKEN = true "
            gdBase.Execute sSQL, dbFailOnError
            gbDSDRUCKEN = True
        Else
            sSQL = "Update KASSEIN Set DSDRUCKEN = false "
            gdBase.Execute sSQL, dbFailOnError
            gbDSDRUCKEN = False
        End If
        
        If Option2(0).Value = True Then
            sSQL = "Update KASSEIN Set DSKLEIN = true "
            gdBase.Execute sSQL, dbFailOnError
            gbDSKLEIN = True
        Else
            sSQL = "Update KASSEIN Set DSKLEIN = false "
            gdBase.Execute sSQL, dbFailOnError
            gbDSKLEIN = False
        End If
        
        If Check3.Value = vbChecked Then
            sSQL = "Update KASSEIN Set DSMeldungErfolg = true "
            gdBase.Execute sSQL, dbFailOnError
            gbDSMeldungErfolg = True
        Else
            sSQL = "Update KASSEIN Set DSMeldungErfolg = false "
            gdBase.Execute sSQL, dbFailOnError
            gbDSMeldungErfolg = False
        End If
        
        
        
        
        If Check4.Value = vbChecked Then
            sSQL = "Update KASSEIN Set DS_GEB_DRUCKEN = true "
            gdBase.Execute sSQL, dbFailOnError
            gbDS_GEB_DRUCKEN = True
        Else
            sSQL = "Update KASSEIN Set DS_GEB_DRUCKEN = false "
            gdBase.Execute sSQL, dbFailOnError
            gbDS_GEB_DRUCKEN = False
        End If
        
        
        
        
        

       Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_Datenschutz"
    Fehler.gsFehlertext = "Im Programmteil Datenschutz ist ein Fehler aufgetreten."
    End Sub





