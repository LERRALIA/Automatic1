VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL165 
   Caption         =   "Kalkulierte Nettospannen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL165.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9945
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   3
         Left            =   10080
         TabIndex        =   18
         Top             =   6960
         Width           =   1695
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   840
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   840
         MaxLength       =   100
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   840
         MaxLength       =   100
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   4200
         Width           =   1215
      End
      Begin sevCommand3.Command Command4 
         Height          =   345
         Index           =   1
         Left            =   11400
         TabIndex        =   7
         ToolTipText     =   "Hilfe"
         Top             =   240
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   0
         Left            =   10080
         TabIndex        =   3
         Top             =   7440
         Width           =   1695
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   2
         Left            =   10080
         TabIndex        =   4
         Top             =   7920
         Width           =   1695
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
         Caption         =   "Schlieﬂen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Kalkulierte Nettospannen aller Artikel mit Bestand sortiert nach Nettospanne anzeigen"
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
         Left            =   120
         TabIndex        =   19
         Top             =   6960
         Width           =   8055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "% Nettospanne"
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
         Index           =   10
         Left            =   2160
         TabIndex        =   17
         Top             =   4320
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "unter"
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
         Left            =   120
         TabIndex        =   16
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "'-' auf dem Etikett"
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
         Index           =   8
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "% Nettospanne"
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
         Index           =   7
         Left            =   2160
         TabIndex        =   14
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "ab"
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
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "'~' auf dem Etikett"
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
         TabIndex        =   12
         Top             =   2880
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "% Nettospanne"
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
         Index           =   5
         Left            =   2160
         TabIndex        =   11
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "ab"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "'+' auf dem Etikett"
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label lblanzeige 
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   8
         Top             =   7920
         Width           =   9255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "Kalkulierte Nettospannen"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   10215
      End
   End
End
Attribute VB_Name = "frmWKL165"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub hinzufuegen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPlus As String
    Dim cNull As String
    Dim cMinus As String
    
    sSQL = "Delete from ETINS"
    gdBase.Execute sSQL, dbFailOnError
    
    If Text3(0).Text <> "" Then
        If IsNumeric(Text3(0).Text) Then
            cPlus = Text3(0).Text
        Else
            anzeige "rot", "Bitte einen Wert eingeben!", lblanzeige
            Text3(0).SetFocus
            Exit Sub
        End If
    Else
        anzeige "rot", "Bitte einen Wert eingeben!", lblanzeige
        Text3(0).SetFocus
        Exit Sub
    End If
    
    If Text3(1).Text <> "" Then
        If IsNumeric(Text3(1).Text) Then
            cNull = Text3(1).Text
        Else
            anzeige "rot", "Bitte einen Wert eingeben!", lblanzeige
            Text3(1).SetFocus
            Exit Sub
        End If
    Else
        anzeige "rot", "Bitte einen Wert eingeben!", lblanzeige
        Text3(1).SetFocus
        Exit Sub
    End If
    
    If Text3(2).Text <> "" Then
        If IsNumeric(Text3(2).Text) Then
            cMinus = Text3(2).Text
        Else
            anzeige "rot", "Bitte einen Wert eingeben!", lblanzeige
            Text3(2).SetFocus
            Exit Sub
        End If
    Else
        anzeige "rot", "Bitte einen Wert eingeben!", lblanzeige
        Text3(2).SetFocus
        Exit Sub
    End If
    
    sSQL = "INSERT into ETINS (ePlus,eNull,eMinus) values "
    sSQL = sSQL & " ('" & cPlus & "','" & cNull & "','" & cMinus & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "erfolgreich gespeichert", lblanzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "hinzufuegen"
    Fehler.gsFehlertext = "Im Programmteil Kalkulierte Nettospannen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub anzeigenEtiNs()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    If NewTableSuchenDBKombi("ETINS", gdBase) = False Then
        CreateTableT2 "ETINS", gdBase
    End If
    
    sSQL = "Select * from ETINS "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!ePLUS) Then
            Text3(0).Text = rsrs!ePLUS
        Else
            Text3(0).Text = ""
        End If
        
        If Not IsNull(rsrs!eNull) Then
            Text3(1).Text = rsrs!eNull
        Else
            Text3(1).Text = ""
        End If
        
        If Not IsNull(rsrs!eMinus) Then
            Text3(2).Text = rsrs!eMinus
        Else
            Text3(2).Text = ""
        End If
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "anzeigenInterart"
    Fehler.gsFehlertext = "Im Programmteil Kalkulierte Nettospannen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Select Case Index
    Case 0
        hinzufuegen
    Case 1
        gsHelpstring = "Kalkulierte Nettospannen"
        frmWKL110.Show 1
    Case 2
        Unload frmWKL165
    Case 3
        negSpanne
End Select
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Kalkulierte Nettospannen fehlt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub negSpanne()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden ermittelt, bitte warten...", lblanzeige
    
    loeschNEW "NEGSPANNE", gdBase
    CreateTableT2 "NEGSPANNE", gdBase
    
    cSQL = "Insert into NEGSPANNE Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , BESTAND "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , MWST "
    cSQL = cSQL & " , EKPR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where Bestand > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update NEGSPANNE set KSPANNE = ((((KVKPR1/(100 + " & gdMWStV & "))* 100) - (EKPR))* 100) / ((KVKPR1/(100 + " & gdMWStV & "))* 100)"
    cSQL = cSQL & " where MWST = 'V' "
    cSQL = cSQL & " and KVKPR1 <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update NEGSPANNE set KSPANNE = ((((KVKPR1/(100 + " & gdMWStE & "))* 100) - (EKPR))* 100) / ((KVKPR1/(100 + " & gdMWStE & "))* 100)"
    cSQL = cSQL & " where MWST = 'E' "
    cSQL = cSQL & " and KVKPR1 <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update NEGSPANNE set KSPANNE = ((((KVKPR1/(100 + " & gdMWStO & "))* 100) - (EKPR))* 100) / ((KVKPR1/(100 + " & gdMWStO & "))* 100)"
    cSQL = cSQL & " where MWST = 'O' "
    cSQL = cSQL & " and KVKPR1 <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    anzeige "normal", "Druckvorschau wird angezeigt...", lblanzeige

    reportbildschirm "WKL165", "aWKL165"
    
    anzeige "normal", "", lblanzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "negSpanne"
    Fehler.gsFehlertext = "Im Programmteil Kalkulierte Nettospannen fehlt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True:
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    Text3(0).Text = ""
    Text3(1).Text = ""
    Text3(2).Text = ""
    
    anzeigenEtiNs
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kalkulierte Nettospannen ist ein Fehler aufgetreten."
    
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



Private Sub Text3_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = glSelBack1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kalkulierte Nettospannen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyEscape Then
        Command4_Click 2
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kalkulierte Nettospannen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kalkulierte Nettospannen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


