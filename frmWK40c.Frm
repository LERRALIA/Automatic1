VERSION 5.00
Begin VB.Form frmWK40c 
   Caption         =   "Lieferantenliste"
   ClientHeight    =   8595
   ClientLeft      =   1635
   ClientTop       =   1770
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWK40c.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
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
      Left            =   5520
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   7680
      Width           =   1095
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Zurück"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Width           =   2055
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Drucken"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   9720
      TabIndex        =   1
      Top             =   7560
      Width           =   2055
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
      Height          =   6150
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11655
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferantenliste"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label1 
      Caption         =   "ausgewähltes Jahr:"
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
      Left            =   5280
      TabIndex        =   3
      Top             =   7440
      Width           =   1695
   End
End
Attribute VB_Name = "frmWK40c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "DRU_TEXT", gdBase
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
Private Sub DruckeListeWK40c()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cpfad As String
    Dim cLBSatz As String
    Dim iFileNr As Integer
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim rsrs As Recordset
    Dim iModus As Integer
    Dim iLiNr As Long
    
    cpfad = gcDBPfad
    If Right(cpfad, 1) <> "\" Then
        cpfad = cpfad & "\"
    End If
    
    iLiNr = Val(Left((frmWKL42!List2.list(frmWKL42!List2.ListIndex)), 6))
    iModus = Val(frmWKL42!Label1(3).Caption)
    
    loeschNEW "DRU_TEXT", gdBase
    
    cSQL = "Create Table DRU_TEXT (LFDNR Long, LINR Long, SATZ Text(190))"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from DRU_TEXT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    lAnzSatz = List2.ListCount
    
    For lAktSatz = 1 To lAnzSatz - 1
        cLBSatz = List2.list(lAktSatz)
        rsrs.AddNew
        rsrs!LFDNR = iModus
        rsrs!linr = iLiNr
        rsrs!SATZ = cLBSatz
        rsrs.Update
    Next lAktSatz
    
    rsrs.Close: Set rsrs = Nothing
    
    reportbildschirm "WKL026", "aWKL40ca"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeListeWK40c"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub LeseDatenWK40c()
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim ctmp As String
    Dim cSQL As String
    Dim cFeld As String
    Dim cLBSatz As String
    Dim rsrs As Recordset
    Dim dWert As Double
    
    Dim clinr As String
    Dim lJahr As Long
    Dim lArtNrMerker As Long
    Dim lJahrMerker As Long
    Dim lMonMerker As Long
    Dim lMon As Long
    Dim dGesamt As Double
    Dim iModus As Integer
    
    ReDim acLbSatz(1 To 16) As String
    ReDim adSumme(1 To 13) As Double


    cLBSatz = frmWKL42!List2.list(frmWKL42!List2.ListIndex)
    
    If cLBSatz = "" Then
        Exit Sub
    End If
    
    cLBSatz = Left(cLBSatz, 6)
    
    clinr = Trim$(cLBSatz)
    cLBSatz = ""
    
    '2 = Umsätze / 3 = Stückzahlen
    iModus = Val(frmWKL42!Label1(3).Caption)
        
    If iModus = 2 Then
        acLbSatz(1) = Space$(6)
        acLbSatz(2) = Space$(35)
        acLbSatz(3) = Space$(4)
        acLbSatz(4) = "     0,00"
        acLbSatz(5) = "     0,00"
        acLbSatz(6) = "     0,00"
        acLbSatz(7) = "     0,00"
        acLbSatz(8) = "     0,00"
        acLbSatz(9) = "     0,00"
        acLbSatz(10) = "     0,00"
        acLbSatz(11) = "     0,00"
        acLbSatz(12) = "     0,00"
        acLbSatz(13) = "     0,00"
        acLbSatz(14) = "     0,00"
        acLbSatz(15) = "     0,00"
        acLbSatz(16) = "     0,00"
    Else
        acLbSatz(1) = Space$(6)
        acLbSatz(2) = Space$(35)
        acLbSatz(3) = Space$(4)
        acLbSatz(4) = "        0"
        acLbSatz(5) = "        0"
        acLbSatz(6) = "        0"
        acLbSatz(7) = "        0"
        acLbSatz(8) = "        0"
        acLbSatz(9) = "        0"
        acLbSatz(10) = "        0"
        acLbSatz(11) = "        0"
        acLbSatz(12) = "        0"
        acLbSatz(13) = "        0"
        acLbSatz(14) = "        0"
        acLbSatz(15) = "        0"
        acLbSatz(16) = "        0"
    End If
    lJahr = Val(Text1.Text)
    
    List2.Clear
    List2.AddItem "ArtNr. Artikelbezeichnung                  Jahr       Jan       Feb       Mrz       Apr       Mai       Jun       Jul       Aug       Sep       Okt       Nov       Dez      Summe"
    
    If iModus = 2 Then
        cSQL = "Select A.ARTNR, B.BEZEICH, A.JAHR, A.MONAT, SUM(A.UMSATZ) as UMSMON, B.LPZ "
    Else
        cSQL = "Select A.ARTNR, B.BEZEICH, A.JAHR, A.MONAT, SUM(A.ANZAHL) as UMSMON, B.LPZ "
    End If
    cSQL = cSQL & "from UMS_ART A, ARTIKEL B "
    cSQL = cSQL & "where B.LINR = " & clinr & " "
    cSQL = cSQL & "and A.ARTNR = B.ARTNR "
    cSQL = cSQL & "and A.JAHR = " & Trim$(Str$(lJahr)) & " "
    cSQL = cSQL & "group by B.LINR, A.ARTNR, B.BEZEICH, A.JAHR, A.MONAT, b.LPZ "
    cSQL = cSQL & "order by B.LPZ, A.ARTNR, A.MONAT "
    
    'MsgBox cSQL
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        lArtNrMerker = -1
        lJahrMerker = -1
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            If cFeld = "300101" Then
                cFeld = cFeld
            End If
            If lArtNrMerker <> Val(cFeld) Then
                If lArtNrMerker <> -1 Then
                    cLBSatz = ""
                    For lcount = 1 To 15
                        cLBSatz = cLBSatz & acLbSatz(lcount) & " "
                    Next lcount
                    If iModus = 2 Then
                        ctmp = Format$(dGesamt, "#####0.00")
                    Else
                        ctmp = Format$(dGesamt, "########0")
                    End If
                    
                    adSumme(13) = adSumme(13) + dGesamt
                    
                    ctmp = Space$(10 - Len(ctmp)) & ctmp
                    cLBSatz = cLBSatz & ctmp & "   "
                    List2.AddItem cLBSatz
                    
                    ReDim acLbSatz(1 To 16) As String
                    If iModus = 2 Then
                        acLbSatz(1) = Space$(6)
                        acLbSatz(2) = Space$(35)
                        acLbSatz(3) = Space$(4)
                        acLbSatz(4) = "     0,00"
                        acLbSatz(5) = "     0,00"
                        acLbSatz(6) = "     0,00"
                        acLbSatz(7) = "     0,00"
                        acLbSatz(8) = "     0,00"
                        acLbSatz(9) = "     0,00"
                        acLbSatz(10) = "     0,00"
                        acLbSatz(11) = "     0,00"
                        acLbSatz(12) = "     0,00"
                        acLbSatz(13) = "     0,00"
                        acLbSatz(14) = "     0,00"
                        acLbSatz(15) = "     0,00"
                        acLbSatz(16) = "     0,00"
                    Else
                        acLbSatz(1) = Space$(6)
                        acLbSatz(2) = Space$(35)
                        acLbSatz(3) = Space$(4)
                        acLbSatz(4) = "        0"
                        acLbSatz(5) = "        0"
                        acLbSatz(6) = "        0"
                        acLbSatz(7) = "        0"
                        acLbSatz(8) = "        0"
                        acLbSatz(9) = "        0"
                        acLbSatz(10) = "        0"
                        acLbSatz(11) = "        0"
                        acLbSatz(12) = "        0"
                        acLbSatz(13) = "        0"
                        acLbSatz(14) = "        0"
                        acLbSatz(15) = "        0"
                        acLbSatz(16) = "        0"
                    End If
                End If
                dGesamt = 0
                lArtNrMerker = Val(cFeld)
                cFeld = Space$(6 - Len(cFeld)) & cFeld
                acLbSatz(1) = cFeld
                
                If Not IsNull(rsrs!BEZEICH) Then
                    cFeld = rsrs!BEZEICH
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cFeld = cFeld & Space$(35 - Len(cFeld))
                acLbSatz(2) = cFeld
                
                If Not IsNull(rsrs!jahr) Then
                    dWert = rsrs!jahr
                Else
                    dWert = 0
                End If
                cFeld = Format$(dWert, "###0")
                cFeld = Space$(4 - Len(cFeld)) & cFeld
                
                acLbSatz(3) = cFeld
                
            End If
            
            If lArtNrMerker = 300079 Then
                lArtNrMerker = lArtNrMerker
            End If
            
            If Not IsNull(rsrs!Monat) Then
                lMon = rsrs!Monat
            Else
                lMon = 1
            End If
            
            If Not IsNull(rsrs!UMSMON) Then
                dWert = rsrs!UMSMON
            Else
                dWert = 0
            End If
            dGesamt = dGesamt + dWert
            
            adSumme(lMon) = adSumme(lMon) + dWert
            
            If iModus = 2 Then
                cFeld = Format$(dWert, "#####0.00")
            Else
                cFeld = Format$(dWert, "########0")
            End If
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            
            acLbSatz(lMon + 3) = cFeld
                
        
            rsrs.MoveNext
        Loop
        
        'Nachlauf
        cLBSatz = ""
        If acLbSatz(1) <> "" Then
            For lcount = 1 To 15
                cLBSatz = cLBSatz & acLbSatz(lcount) & " "
            Next lcount
            
            adSumme(13) = adSumme(13) + dGesamt
            
            If iModus = 2 Then
                ctmp = Format$(dGesamt, "#####0.00")
            Else
                ctmp = Format$(dGesamt, "########0")
            End If
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            cLBSatz = cLBSatz & ctmp & "   "
            List2.AddItem cLBSatz

        End If
        cLBSatz = ""
        ReDim acLbSatz(1 To 16) As String
        If iModus = 2 Then
            acLbSatz(1) = Space$(6)
            acLbSatz(2) = Space$(35)
            acLbSatz(3) = Space$(4)
            acLbSatz(4) = "     0,00"
            acLbSatz(5) = "     0,00"
            acLbSatz(6) = "     0,00"
            acLbSatz(7) = "     0,00"
            acLbSatz(8) = "     0,00"
            acLbSatz(9) = "     0,00"
            acLbSatz(10) = "     0,00"
            acLbSatz(11) = "     0,00"
            acLbSatz(12) = "     0,00"
            acLbSatz(13) = "     0,00"
            acLbSatz(14) = "     0,00"
            acLbSatz(15) = "     0,00"
            acLbSatz(16) = "     0,00"
        Else
            acLbSatz(1) = Space$(6)
            acLbSatz(2) = Space$(35)
            acLbSatz(3) = Space$(4)
            acLbSatz(4) = "        0"
            acLbSatz(5) = "        0"
            acLbSatz(6) = "        0"
            acLbSatz(7) = "        0"
            acLbSatz(8) = "        0"
            acLbSatz(9) = "        0"
            acLbSatz(10) = "        0"
            acLbSatz(11) = "        0"
            acLbSatz(12) = "        0"
            acLbSatz(13) = "        0"
            acLbSatz(14) = "        0"
            acLbSatz(15) = "        0"
            acLbSatz(16) = "        0"
        End If
        acLbSatz(2) = "Gesamtsumme:" & Space$(23)
        acLbSatz(3) = Format$(lJahr, "###0")
                
        For lcount = 1 To 12
            If iModus = 2 Then
                 acLbSatz(lcount + 3) = Format$(adSumme(lcount), "#####0.00")
            Else
                 acLbSatz(lcount + 3) = Format$(adSumme(lcount), "########0")
            End If
        Next lcount
        
        cLBSatz = cLBSatz & acLbSatz(1) & " "
        cLBSatz = cLBSatz & acLbSatz(2) & " "
        cLBSatz = cLBSatz & acLbSatz(3) & " "
        
        For lcount = 4 To 15
            ctmp = acLbSatz(lcount)
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            cLBSatz = cLBSatz & ctmp & " "
        Next lcount
        If iModus = 2 Then
            ctmp = Format$(adSumme(13), "#####0.00")
        Else
            ctmp = Format$(adSumme(13), "########0")
        End If
        ctmp = Space$(9 - Len(ctmp)) & ctmp
        cLBSatz = cLBSatz & ctmp & "   "
        List2.AddItem cLBSatz
        
    End If
    
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseDatenWK40c"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            DruckeListeWK40c
            
        Case Is = 1
            Unload frmWK40c
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing

    Text1.Text = Year(Now)

    LeseDatenWK40c
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3375 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Form_Load"
        Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub List2_Click()
    On Error GoTo LOKAL_ERROR
    
    If List2.ListCount > 1 Then
        If List2.ListIndex = 0 Then
            List2.Selected(List2.ListIndex) = False
            List2.ListIndex = 1
            List2.Selected(List2.ListIndex) = True
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change()
    On Error GoTo LOKAL_ERROR
    
    If Len(Text1) = 4 Then
        LeseDatenWK40c
    End If
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
