VERSION 5.00
Begin VB.Form frmWKL118 
   Caption         =   "Kundenbestellungen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   11535
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Bestellen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   8520
      TabIndex        =   4
      Top             =   7080
      Width           =   3135
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "nicht erneut bestellen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   8520
      TabIndex        =   1
      Top             =   7800
      Width           =   3135
   End
   Begin sevCommand3.Command Command5 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "?"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   11535
   End
   Begin VB.Label Label1 
      Caption         =   "Wollen Sie den/diese Artikel wirklich noch einmal bestellen?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   9375
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kundenbestellungen"
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
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   9495
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
   Begin VB.Label lbl1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   7935
   End
End
Attribute VB_Name = "frmWKL118"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            gb118bestell = True
            Unload frmWKL118
        Case 3
            gb118bestell = False
            Unload frmWKL118
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Select Case Index
    
        Case 11
            gsHelpstring = "Kundenbestellungen"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift

    LogtoStart Me
    
    If NewTableSuchenDBKombi("kundberr", gdBase) Then
    
        zeigKundBerr
        
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigKundBerr()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim rsrs1   As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    Dim cArtNr  As String
    Dim cAm     As String
    Dim cUm     As String
    Dim lMaxi   As Long
    Dim lbed    As Long
    
    List1.Clear
    List2.Clear
    List2.AddItem "ArtNr  Bezeichnung                                BM/14T zuletzt am   um     Menge   Bed"
    Screen.MousePointer = 11
    
    loeschNEW "KBERRS", gdBase
    CreateTable "KBERRS", gdBase
    
    sSQL = "insert into KBERRS select sum(Bestelltmenge) as menge " ',max(Bestelltam) as ZULAM,max(Bestelltum) as zulum "
    sSQL = sSQL & ", artnr"
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", EKPR "
    sSQL = sSQL & ", VKPR "
    sSQL = sSQL & ", MWST "
    sSQL = sSQL & " from kundberr group by "
    sSQL = sSQL & " artnr"
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", EKPR "
    sSQL = sSQL & ", VKPR "
    sSQL = sSQL & ", MWST "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("KBERRS")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            Else
                cArtNr = ""
            End If
            
            sSQL = "select max(bestelltam) as maxi from KUNDBEST where KUNDBEST.Artnr = " & cArtNr
            Set rsrs1 = gdBase.OpenRecordset(sSQL)
            If Not rsrs1.EOF Then
                If Not IsNull(rsrs1!maxi) Then
                    cAm = rsrs1!maxi
                End If
            End If
            rsrs1.Close: Set rsrs1 = Nothing
            
            sSQL = "Update KBERRS "
            sSQL = sSQL & " set KBERRS.ZULAM = " & CLng(DateValue(cAm))
            sSQL = sSQL & " where KBERRS.artnr = " & cArtNr
            gdBase.Execute sSQL, dbFailOnError
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Set rsrs = gdBase.OpenRecordset("KBERRS")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            Else
                cArtNr = ""
            End If
            
            
            If Not IsNull(rsrs!zulam) Then
                cAm = rsrs!zulam
            Else
                cAm = ""
            End If
            
            sSQL = "select max(bestelltum) as maxi from KUNDBEST where KUNDBEST.Artnr = " & cArtNr
            sSQL = sSQL & " and KUNDBEST.BESTELLTAM = " & CLng(DateValue(cAm))
            
            Set rsrs1 = gdBase.OpenRecordset(sSQL)
            If Not rsrs1.EOF Then
                If Not IsNull(rsrs1!maxi) Then
                    cUm = rsrs1!maxi
                
                End If
            End If
            rsrs1.Close: Set rsrs1 = Nothing
            
            sSQL = "Update KBERRS "
            sSQL = sSQL & " set KBERRS.ZULum = '" & cUm & "'"
            sSQL = sSQL & " where KBERRS.artnr = " & cArtNr
            gdBase.Execute sSQL, dbFailOnError
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Set rsrs = gdBase.OpenRecordset("KBERRS")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            Else
                cArtNr = ""
            End If
            
            If Not IsNull(rsrs!zulum) Then
                cUm = rsrs!zulum
            Else
                cUm = ""
            End If
            
            If Not IsNull(rsrs!zulam) Then
                cAm = rsrs!zulam
            Else
                cAm = ""
            End If
            
            sSQL = "select sum(bestelltmenge) as maxi, max(bednu) as mbed from KUNDBEST where KUNDBEST.Artnr = " & cArtNr
            sSQL = sSQL & " and KUNDBEST.BESTELLTAM = " & CLng(DateValue(cAm))
            sSQL = sSQL & " and KUNDBEST.BESTELLTUM = '" & cUm & "'"
            
            Set rsrs1 = gdBase.OpenRecordset(sSQL)
            If Not rsrs1.EOF Then
                If Not IsNull(rsrs1!maxi) Then
                    lMaxi = rsrs1!maxi
                Else
                    lMaxi = 0
                End If
                
                If Not IsNull(rsrs1!mbed) Then
                    lbed = rsrs1!mbed
                Else
                    lbed = 0
                End If
            End If
            rsrs1.Close: Set rsrs1 = Nothing
            
            sSQL = "Update KBERRS "
            sSQL = sSQL & " set KBERRS.ZULMENGE = " & lMaxi
            sSQL = sSQL & " , KBERRS.BEDNU = " & lbed
            sSQL = sSQL & " where KBERRS.artnr = " & cArtNr
            gdBase.Execute sSQL, dbFailOnError
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Set rsrs = gdBase.OpenRecordset("KBERRS")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(6 - Len(cFeld))
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(40 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & "   "
            
            If Not IsNull(rsrs!menge) Then
                cFeld = rsrs!menge
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(6 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!zulam) Then
                cFeld = rsrs!zulam
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(12 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!zulum) Then
                cFeld = rsrs!zulum
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(10 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!ZULMENGE) Then
                cFeld = rsrs!ZULMENGE
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(4 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            
            If Not IsNull(rsrs!BEDNU) Then
                cFeld = rsrs!BEDNU
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(4 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            List1.AddItem cLBSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "KBERRS", gdBase
    loeschNEW "KUNDBERR", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigKundBerr"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub





