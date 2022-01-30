VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL92 
   Caption         =   "Kundenbestellungen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL92.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox Combo1 
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
      Left            =   120
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   7800
      Width           =   2895
   End
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   16
      Top             =   7800
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
      Caption         =   "Ändern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   2
      Left            =   5280
      TabIndex        =   15
      Top             =   7800
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
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   1
      Left            =   7440
      TabIndex        =   14
      Top             =   7800
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
      Caption         =   "Drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   12
      Top             =   7800
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
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame5"
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   11775
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4785
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   1
         Top             =   1080
         Width           =   11535
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   11535
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame3"
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   6120
         Width           =   11535
         Begin VB.OptionButton Option4 
            Caption         =   "Bestelldatum"
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
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Tag             =   "bestelltam asc"
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Bediener"
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
            Index           =   4
            Left            =   2280
            TabIndex        =   7
            Tag             =   "Bednu"
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Filiale"
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
            Index           =   5
            Left            =   3960
            TabIndex        =   6
            Tag             =   "Filiale"
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Menge"
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
            Index           =   7
            Left            =   5400
            TabIndex        =   5
            Tag             =   "bestelltmenge desc"
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Filiale Datum"
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
            Index           =   8
            Left            =   6960
            TabIndex        =   4
            Tag             =   "Filiale , bestelltam asc"
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Artikelstatus"
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
            Index           =   0
            Left            =   9360
            TabIndex        =   3
            Tag             =   "StatusARTIKEL"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Sortierung nach"
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
            Index           =   7
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Label lblanzeige 
         BackColor       =   &H00C0C000&
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Verteilte Artikel, die zur Übertragung bereitstehen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   6735
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kundenbestellungen"
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
      TabIndex        =   13
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmWKL92"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    
    With Frame5
        .Height = 6855
        .Left = 0
        .Top = 840
        .Width = 11775
        .BorderStyle = 0
        
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil verteilte Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
    Case 0
        Unload frmWKL92
    Case 1
        drucken
    Case 2
        loescheausKUNDBEST
    Case 3
        UpdateKUNDBEST
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub loescheausKUNDBEST()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz         As String
    Dim cArtNr          As String
    Dim cBestelltam     As String
    Dim cBestelltum     As String
    Dim cSQL            As String
    
    If List3.ListIndex < 0 Then
        MsgBox "Bitte einen Eintrag auswählen!", vbInformation, "Winkiss Hinweis:"
        List3.SetFocus
        Exit Sub
    End If
    
    Dim bFound As Boolean
    Dim lCount As Long
    

    bFound = False
    
    For lCount = 0 To List3.ListCount - 1
        If List3.Selected(lCount) = True Then
            bFound = True
            
            cLBSatz = List3.list(lCount)
            cBestelltam = Left$(cLBSatz, 8)
            cBestelltum = Mid$(cLBSatz, 10, 8)
            cArtNr = Mid$(cLBSatz, 28, 6)
                
            cSQL = "Delete from KUNDBEST where KUNDNR = " & gckundnr
            cSQL = cSQL & " and ARTNR = " & cArtNr
            cSQL = cSQL & " and BESTELLTAM = " & CLng(DateValue(cBestelltam))
            cSQL = cSQL & " and BESTELLTUM = '" & cBestelltum & "'"
            gdBase.Execute cSQL, dbFailOnError
        End If
    Next lCount
    If Not bFound Then
        MsgBox "Zum Löschen bitte mindestens einen Listeneintrag auswählen!", vbInformation, "Winkiss Hinweis:"
        List3.SetFocus
        Exit Sub
    End If
        
    ZeigArtHistInList "KUBE", List3, gckundnr, "StatusARTIKEL"
    anzeige "normal", gckundnr, lblanzeige
'    gckundnr = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loescheausKUNDBEST"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub UpdateKUNDBEST()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz         As String
    Dim cArtNr          As String
    Dim cBestelltam     As String
    Dim cBestelltum     As String
    Dim cSQL            As String
    Dim cSTATUS         As String
    
    
    If List3.ListIndex < 0 Then
        MsgBox "Bitte einen Eintrag auswählen!", vbInformation, "Winkiss Hinweis:"
        List3.SetFocus
        Exit Sub
    End If
    
     Select Case Combo1.Text
        Case "noch nicht bestellt"
            cSTATUS = "INBESTELLUNG"
        Case "ist bestellt"
            cSTATUS = "BESTELLT"
        Case "geliefert"
            cSTATUS = "GELIEFERT"
        Case "nicht geliefert"
            cSTATUS = "NICHTGELIEFERT"
        Case Else
            MsgBox "Bitte einen Eintrag auswählen!", vbInformation, "Winkiss Hinweis:"
            Combo1.SetFocus
            Exit Sub
        
    End Select
            
    Dim bFound As Boolean
    Dim lCount As Long
    
    bFound = False
    
    For lCount = 0 To List3.ListCount - 1
        If List3.Selected(lCount) = True Then
            bFound = True
            
            cLBSatz = List3.list(lCount)
            cBestelltam = Left$(cLBSatz, 8)
            cBestelltum = Mid$(cLBSatz, 10, 8)
            cArtNr = Mid$(cLBSatz, 28, 6)
                
            cSQL = "Update KUNDBEST set STATUSARTIKEL = '" & cSTATUS & "'"
            cSQL = cSQL & " Where KUNDNR = " & gckundnr
            cSQL = cSQL & " and ARTNR = " & cArtNr
            cSQL = cSQL & " and BESTELLTAM = " & CLng(DateValue(cBestelltam))
            cSQL = cSQL & " and BESTELLTUM = '" & cBestelltum & "'"
            gdBase.Execute cSQL, dbFailOnError
            
        End If
    Next lCount
    If Not bFound Then
        MsgBox "Zum Ändern bitte mindestens einen Listeneintrag auswählen!", vbInformation, "Winkiss Hinweis:"
        List3.SetFocus
        Exit Sub
    End If
    
    
        
       
    ZeigArtHistInList "KUBE", List3, gckundnr, "StatusARTIKEL"
    anzeige "normal", gckundnr, lblanzeige
'    gckundnr = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateKUNDBEST"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren
    
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    List1.AddItem "Bedatum  Uhrzeit     Menge Artnr  Artikelbezeichnung                 Fil Preis  Bed  Artikelstatus"
    
    
    ZeigArtHistInList "KUBE", List3, gckundnr, "StatusARTIKEL"
    anzeige "normal", gckundnr, lblanzeige
'    gckundnr = ""
    
    fuellecombo1

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
   
End Sub
Private Sub fuellecombo1()
    On Error GoTo LOKAL_ERROR
    
    Combo1.Clear
    Combo1.AddItem "noch nicht bestellt"
    Combo1.AddItem "ist bestellt"
    Combo1.AddItem "geliefert"
    Combo1.AddItem "nicht geliefert"
    Combo1.AddItem "bitte auswählen"
    
    Combo1.Text = "bitte auswählen"
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo1"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
    
End Sub
Private Sub drucken()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String

    
    loeschNEW "KUOKB", gdBase
    CreateTable "KUOKB", gdBase
    
    sSQL = "Insert into KUOKB Select "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", BEZEICH"
    sSQL = sSQL & ", BEDNU  "
    sSQL = sSQL & ", Filiale "
    sSQL = sSQL & ", STATUSARTIKEL "
    sSQL = sSQL & ", BESTELLTAM  "
    sSQL = sSQL & ", BESTELLTUM  "
    sSQL = sSQL & ", BESTELLTPREIS  "
    sSQL = sSQL & ", BESTELLTMENGE  "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & "  from KUBE "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUOKB SET StatusARTIKEL = 'noch nicht bestellt'  "
    sSQL = sSQL & " where StatusARTIKEL = 'A' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUOKB SET StatusARTIKEL = 'ist bestellt'  "
    sSQL = sSQL & " where StatusARTIKEL = 'B' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUOKB SET StatusARTIKEL = 'geliefert'  "
    sSQL = sSQL & " where StatusARTIKEL = 'C' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUOKB SET StatusARTIKEL = 'nicht geliefert'  "
    sSQL = sSQL & " where StatusARTIKEL = 'D' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUOKB inner join KUNDEN on KUOKB.KUNDNR = KUNDEN.KUNDNR"
    sSQL = sSQL & " SET KUOKB.TEL = KUNDEN.TEL "
    sSQL = sSQL & ", KUOKB.FAXNR = KUNDEN.FAXNR "
    sSQL = sSQL & ", KUOKB.EMAIL = KUNDEN.EMAIL "
    sSQL = sSQL & ", KUOKB.MOBILTEL = KUNDEN.MOBILTEL "
    sSQL = sSQL & ", KUOKB.VORNAME = KUNDEN.VORNAME "
    
    sSQL = sSQL & ", KUOKB.NAME = KUNDEN.NAME "
    sSQL = sSQL & ", KUOKB.STRASSE = KUNDEN.STRASSE "
    sSQL = sSQL & ", KUOKB.PLZ = KUNDEN.PLZ "
    sSQL = sSQL & ", KUOKB.ORT = KUNDEN.STADT "
    sSQL = sSQL & ", KUOKB.TITEL = KUNDEN.TITEL "
    sSQL = sSQL & ", KUOKB.FIRMA = KUNDEN.FIRMA "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
      
    reportbildschirm "", "aZEN76"
    
'    loeschNEW "KUOKB", gdBase
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten. "
    
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

Private Sub Option4_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    ZeigArtHistInList "KUBE", List3, gckundnr, Option4(Index).Tag
'    gckundnr = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option4_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



