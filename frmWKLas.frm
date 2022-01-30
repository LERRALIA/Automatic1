VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWKLas 
   BackColor       =   &H00C0C000&
   Caption         =   "Unterschrittene Mindestmenge ermitteln"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   -180
   ClientWidth     =   11880
   Icon            =   "frmWKLas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox ListLief 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   5760
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   4575
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
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   10815
   End
   Begin sevCommand3.Command cmdQ 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "?"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   240
      Width           =   375
   End
   Begin sevCommand3.Command cmdEnd 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Schließen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      TabIndex        =   7
      Top             =   7440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8520
      MaxLength       =   6
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin sevCommand3.Command cmdUMErmitteln 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Erstellen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   7440
      Width           =   3135
   End
   Begin sevCommand3.Command cmdPrint 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Drucken"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   5
      Top             =   7440
      Width           =   3135
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
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   10815
   End
   Begin MSComctlLib.ProgressBar pbrZeit 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferantennummer"
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
      Index           =   2
      Left            =   3000
      TabIndex        =   11
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   " bis:"
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
      Left            =   7320
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "von:"
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
      Left            =   5160
      TabIndex        =   9
      Top             =   360
      Width           =   495
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
      Left            =   480
      TabIndex        =   8
      Top             =   7080
      Width           =   10815
   End
End
Attribute VB_Name = "frmWKLas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function Check() As Integer
    On Error GoTo LOKAL_ERROR
    
    If Text1(0).Text = "" Then
        lblAnzeige.Caption = "Sie müssen den Lieferantenbereich angeben!"
        lblAnzeige.Refresh
        Text1(0).SetFocus
        Check = 1
    Else
        Check = 2
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
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
Private Sub cmdEnd_Click()
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "DRU_TEXT", gdBase
    
    Unload frmWKLas
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEnd_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub DruckeListe()
    On Error GoTo LOKAL_ERROR
    
    reportbildschirm "WKL024a", "aWKL40da"
    

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeListe"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub cmdQ_Click()
    On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    
    
    sSQL = "SELECT LIEFBEZ, LINR From LISRT"
    sSQL = sSQL & " Order BY LIEFBEZ"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!linr) Then
                cFeld = rsrs!linr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(7 - Len(cFeld))
            cLBSatz = cFeld
                
            If Not IsNull(rsrs!LIEFBEZ) Then
                cFeld = rsrs!LIEFBEZ
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld
                
            ListLief.AddItem cLBSatz
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
            
    ListLief.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdQ_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
        
End Sub

Private Sub cmdQ_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    ListLief.Visible = False
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdQ_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub


Private Sub Form_Click()
    On Error GoTo LOKAL_ERROR
    
    ListLief.Visible = False
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub ListLief_Click()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = Left(ListLief.list(ListLief.ListIndex), 6)
    Text1(0).SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ListLief_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub



Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    If Index = 1 And Text1(0).Text = "" Then
        lblAnzeige.Caption = "Sie müssen erst einen Startwert eingeben!"
        lblAnzeige.Refresh
        Text1(0).SetFocus
    ElseIf Text1(0).Text <> "" Then
        lblAnzeige.Caption = ""
        lblAnzeige.Refresh
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = vbWhite
    
    If Index = 0 Then
        Text1(1).Text = Text1(0).Text
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo LOKAL_ERROR
    
    ListLief.Visible = False
    lblAnzeige.Caption = ""
    lblAnzeige.Refresh
    
    If List2.ListCount = 0 Then
        lblAnzeige.Caption = "Kein ausdruckbares Ergebnis vorhanden"
        lblAnzeige.Refresh
    Else
        lblAnzeige.Caption = "Berichtsvorschau wird geladen..."
        lblAnzeige.Refresh
        DruckeListe
        lblAnzeige.Caption = ""
        lblAnzeige.Refresh
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPrint_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdUMErmitteln_Click()
    On Error GoTo LOKAL_ERROR
    
    ListLief.Visible = False
    lblAnzeige.Caption = ""
    lblAnzeige.Refresh

    
    If Check = 2 Then

        Screen.MousePointer = 11

        ListeFuellen
        Screen.MousePointer = 0
    End If
    
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdUMErmitteln_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub ListeFuellen()
On Error GoTo LOKAL_ERROR

    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim sSQL        As String
    Dim lcount      As Long
    Dim counter     As Integer
    Dim counter1    As Integer
    Dim rsrs        As Recordset
    Dim rsrs1       As Recordset
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim cSQL        As String
    
    lVon = CLng(Text1(0))
    lBis = CLng(Text1(1))
    
    lblAnzeige.Caption = "Artikeldaten für diesen Bereich werden ermittelt..."
    lblAnzeige.Refresh
    
    List2.Clear
    
    sSQL = "Select LINR, ARTNR, BEZEICH, BESTAND, MINBEST "
    sSQL = sSQL & " from ARTIKEL "
    sSQL = sSQL & " where BESTAND < MINBEST "
    sSQL = sSQL & " and GEFUEHRT = 'J' "
    sSQL = sSQL & " and LINR BETWEEN " & lVon & " and  " & lBis & ""
    sSQL = sSQL & " order by LINR, BEZEICH"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        pbrZeit.Max = rsrs.RecordCount
        
        
        If rsrs.RecordCount > 1000 Then
            lblAnzeige.Caption = "Es wurden zu viele Datensätze gefunden. Bitte schränken Sie den Bereich weiter ein!"
            lblAnzeige.Refresh
            Text1(0).Text = ""
            Text1(1).SetFocus
        Else
            pbrZeit.Visible = True
    
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                counter = counter + 1
                pbrZeit.Value = counter
            
                If Not IsNull(rsrs!linr) Then
                    cFeld = rsrs!linr
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cFeld = Space$(6 - Len(cFeld)) & cFeld
                cLBSatz = cFeld & " "
                
                If Not IsNull(rsrs!artnr) Then
                    cFeld = rsrs!artnr
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cFeld = Space$(6 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                
                If Not IsNull(rsrs!BEZEICH) Then
                    cFeld = rsrs!BEZEICH
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cFeld = cFeld & Space$(35 - Len(cFeld))
                cLBSatz = cLBSatz & cFeld & " "
                
                If Not IsNull(rsrs!BESTAND) Then
                    lcount = rsrs!BESTAND
                Else
                    lcount = 0
                End If
                cFeld = Trim$(Str$(lcount))
                cFeld = Space$(7 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld
                
                If Not IsNull(rsrs!MINBEST) Then
                    lcount = rsrs!MINBEST
                Else
                    lcount = 0
                End If
                cFeld = Trim$(Str$(lcount))
                cFeld = Space$(14 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld
                
                List2.AddItem cLBSatz
                rsrs.MoveNext
            Loop
            
            lblAnzeige.Caption = "Die ermittelten Datensätze werden in die Liste übertragen..."
            lblAnzeige.Refresh
            

            loesch "DRU_TEXT"
            
            cSQL = "Create Table DRU_TEXT (DRUCK Text(80))"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Select * from DRU_TEXT"
            Set rsrs1 = gdBase.OpenRecordset(cSQL)
            
            lAnzSatz = List2.ListCount
            
            pbrZeit.Refresh
            pbrZeit.Max = lAnzSatz
            counter = 0
            
            For lAktSatz = 0 To lAnzSatz - 1
            
                counter = counter + 1
                pbrZeit.Value = counter
                
                cLBSatz = List2.list(lAktSatz)
                rsrs1.AddNew
                rsrs1!DRUCK = cLBSatz
                rsrs1.Update
            Next lAktSatz
            rsrs1.Close: Set rsrs1 = Nothing
            
            pbrZeit.Visible = False
            lblAnzeige.Caption = "Die Ermittlung der Daten ist abgeschlossen."
            lblAnzeige.Refresh
        End If
        
    Else
        lblAnzeige.Caption = "Es wurden keine Datensätze gefunden."
        lblAnzeige.Refresh
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ListeFuellen"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
        
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing

    List1.AddItem "LiefNr ArtNr. Artikelbezeichnung                     Bestand    Mindestmenge"
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Mindestmengenermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

