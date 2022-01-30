VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RettungHistorie 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "FDateienRettung_Historie"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton optName 
      Caption         =   "F_Datei Name: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.OptionButton optDatum 
      Caption         =   "gerettete F_Dateien vom : "
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
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.TextBox txtDname 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "an die Zentrale wieder absenden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   5160
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "suchen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112394241
      CurrentDate     =   44487
   End
   Begin MSFlexGridLib.MSFlexGrid dgvRes 
      Height          =   2325
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4101
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      RowHeightMin    =   300
      AllowUserResizing=   3
   End
   Begin VB.Label Label4 
      Caption         =   "( Format: F0000000 )"
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Historie der F_Dateien"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblKass 
      Caption         =   "*****"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblFil 
      Caption         =   "*****"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Kasse :"
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
      Left            =   5400
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Filiale :"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "RettungHistorie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR

'nach an die Zentrale geschickte geretteten F-Dateien suchen
 
ClearGridView

Dim rsrsR As Recordset

If optDatum.value = True Then

  Set rsrsR = gdBase.OpenRecordset("SELECT * FROM STEUERKI WHERE IsRettungFDatei=true AND KASNUM=" & gcKasNum & " AND Datevalue(RettungZumTag)=CDate('" & DTPicker1.value & "') AND RettungZumTag <> null AND KASNUM <> null AND EndDatname <> null order by DATUM desc")

Else

  Set rsrsR = gdBase.OpenRecordset("SELECT * FROM STEUERKI WHERE DATNAME='" & txtDname.Text & "' order by DATUM desc")
 
End If

If Not rsrsR.EOF Then
   
   rsrsR.MoveFirst
    
    Do While Not rsrsR.EOF
            
      
         
         dgvRes.Rows = dgvRes.Rows + 1
         Dim newRow As Integer
         newRow = dgvRes.Rows - 1
         
         If Not IsNull(rsrsR!lfnr) Then
            dgvRes.TextMatrix(newRow, 1) = rsrsR!lfnr
         End If
         
         If Not IsNull(rsrsR!Datname) Then
            dgvRes.TextMatrix(newRow, 2) = rsrsR!Datname
         End If
         
         If Not IsNull(rsrsR!Datum) Then
            dgvRes.TextMatrix(newRow, 3) = rsrsR!Datum
         End If
         
         If Not IsNull(rsrsR!RettungZumTag) Then
            dgvRes.TextMatrix(newRow, 4) = rsrsR!RettungZumTag
         End If
         
         If Not IsNull(rsrsR!kasnum) Then
            dgvRes.TextMatrix(newRow, 5) = rsrsR!kasnum
         End If
         
         If Not IsNull(rsrsR!EndDatname) Then
            dgvRes.TextMatrix(newRow, 6) = rsrsR!EndDatname
         End If
           
         If Not IsNull(rsrsR!gesendet) Then
            dgvRes.TextMatrix(newRow, 7) = rsrsR!gesendet
         End If
         
         rsrsR.MoveNext
        
    Loop
    dgvRes.Row = 0
End If

If Not dgvRes.Rows > 1 Then
   MsgBox ("keine Ergebnisse gefunden !!!")
End If


Exit Sub
LOKAL_ERROR:
     

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "RettungHistorie"
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub



Private Sub Command3_Click()
On Error GoTo LOKAL_ERROR
     
    'hat dgvRes Datensätze ?
    If Not dgvRes.Rows > 1 Then
     MsgBox ("bitte einen Datensatz auswählen !!!")
     Exit Sub
    End If
    
    'wurde einen Datensatz von dgvRes ausgewählt ?
    If Not dgvRes.Row > 0 Then
     MsgBox ("bitte einen Datensatz auswählen !!!")
     Exit Sub
    End If
     
    
    Dim ausTag As String
    Dim Datnamee As String
    
    ausTag = dgvRes.TextMatrix(dgvRes.Row, 4)
    Datnamee = dgvRes.TextMatrix(dgvRes.Row, 2)
    Datnamee = Datnamee & ".lzh"
    
    If Trim(ausTag) = "" Then
    
     MsgBox ("die ausgewählte Datei war keine gerettete Datei !!!")
     Exit Sub
     
    End If
    
    If Dir(gcDBPfad & "\END Rettung\" & ausTag, vbDirectory) = "" Then
      MsgBox ("Datei nicht gefunden !!!" & vbNewLine & vbNewLine & gcDBPfad & "\END Rettung\" & ausTag & "\" & Datnamee)
      Exit Sub
    End If
    
    If Not FileExists(gcDBPfad & "\END Rettung\" & ausTag & "\" & Datnamee) Then
      MsgBox ("Datei nicht gefunden !!!" & vbNewLine & vbNewLine & gcDBPfad & "\END Rettung\" & ausTag & "\" & Datnamee)
      Exit Sub
    End If
    
    
    
    
    'gerettete F-Datei an die Zentrale absenden <<<<<<<<<<<<<<<<<<<<<<< START
     
         Dim ireslt1 As Integer
         ireslt1 = MsgBox("an die Zentrale absenden ?", vbQuestion + vbYesNo, "gerettete FDatei an die Zentrale")
                       
         If ireslt1 = vbYes Then
         
         
                If Dir(gcDBPfad & "\RettungAnZentrale", vbDirectory) = "" Then
                       MkDir gcDBPfad & "\RettungAnZentrale"
                Else
                       If Not Dir(gcDBPfad & "\RettungAnZentrale\*.*") = "" Then
                        Kill gcDBPfad & "\RettungAnZentrale\*.*"
                       End If
                      
                End If
        
        
                Dim lRet  As Long
                Dim lfail As Long
                 
                lRet = CopyFile(gcDBPfad & "\END Rettung\" & ausTag & "\" & Datnamee, gcDBPfad & "\RettungAnZentrale\" & Datnamee, lfail)
                    
                If Not lRet = 1 Then
                     
                     MsgBox ("Kopieren fehlgeschlagen !!! " & vbNewLine & vbNewLine & "Dateiname: " & Datnamee & vbNewLine & vbNewLine & " von:" & vbNewLine & gcDBPfad & "\END Rettung\" & ausTag & vbNewLine & vbNewLine & "in: " & vbNewLine & gcDBPfad & "\RettungAnZentrale")
                     Exit Sub
                     
                Else
                
gerettete_FDatei_An_Die_Zentrale:

                       geretteteF_DateiErfolgreichAbgeschickt = False
                       giKissFtpMode = 50
                       frmWKL38.Show 1
                       
                       
                       If geretteteF_DateiErfolgreichAbgeschickt Then
                       
                            MsgBox ("erfolgreich an die Zentrale abgeschickt.")
                             
                            If Not UPDATE_STEUERKI Then
                              MsgBox ("das Schreiben in der Tabelle ' STEUERKI ' war fehlgeschlagen !!!")
                            
                            Else
                            'die Suche-Ergebnisse aktualisieren
                              Command1_Click
                            End If
                             
                            Exit Sub
                       Else
                             
                             'wenn gcDBPfad\RettungAnZentrale  nicht leer ist... dann frag mal für nochmal versuchen
                             If Not Dir(gcDBPfad & "\RettungAnZentrale\*.*") = "" Then
                             
                                    Dim ireslt2 As Integer
                                    ireslt2 = MsgBox("fehlgeschlagen !!! " & vbNewLine & vbNewLine & " möchten Sie erneut versuchen ?", vbQuestion + vbYesNo, "gerettete FDatei an die Zentrale")
        
                                    If ireslt2 = vbYes Then
                                       GoTo gerettete_FDatei_An_Die_Zentrale
                                    Else
                                       Exit Sub
                                    End If
                             Else
                                    If Not UPDATE_STEUERKI Then
                                     MsgBox ("das Schreiben in der Tabelle ' STEUERKI ' war fehlgeschlagen !!!")
                                    End If
                                    Exit Sub
                             End If
                              
                       End If
                       
                       
                End If
        
         Else
          Exit Sub
         End If
     
     'gerettete F-Datei an die Zentrale absenden <<<<<<<<<<<<<<<<<<<<<<< ENDE
    
    
  
    

Exit Sub
LOKAL_ERROR:
     

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "RettungHistorie"
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
    
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

 lblFil.Caption = gcFilNr
 lblFil.Refresh
    
 lblKass.Caption = gcKasNum
 lblFil.Refresh
 
 
 dgvRes.ColAlignment(1) = flexAlignLeftCenter
 dgvRes.ColAlignment(2) = flexAlignLeftCenter
 dgvRes.ColAlignment(3) = flexAlignLeftCenter
 dgvRes.ColAlignment(4) = flexAlignLeftCenter
 dgvRes.ColAlignment(5) = flexAlignLeftCenter
 dgvRes.ColAlignment(6) = flexAlignLeftCenter
 dgvRes.ColAlignment(7) = flexAlignLeftCenter
 
 dgvRes.ColWidth(6) = 3500
 dgvRes.ColWidth(3) = 1700
 
 dgvRes.TextMatrix(0, 1) = "LFNR"
 dgvRes.TextMatrix(0, 2) = "DATNAME"
 dgvRes.TextMatrix(0, 3) = "DATUM"
 dgvRes.TextMatrix(0, 4) = "RettungZumTag"
 dgvRes.TextMatrix(0, 5) = "KASNUM"
 dgvRes.TextMatrix(0, 6) = "EndDatname"
 dgvRes.TextMatrix(0, 7) = "gesendet"
 
 
 Exit Sub
LOKAL_ERROR:
     

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "RettungHistorie"
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Sub ClearGridView()
On Error GoTo LOKAL_ERROR
   
 dgvRes.Clear
 dgvRes.Rows = 1

 dgvRes.TextMatrix(0, 1) = "LFNR"
 dgvRes.TextMatrix(0, 2) = "DATNAME"
 dgvRes.TextMatrix(0, 3) = "DATUM"
 dgvRes.TextMatrix(0, 4) = "RettungZumTag"
 dgvRes.TextMatrix(0, 5) = "KASNUM"
 dgvRes.TextMatrix(0, 6) = "EndDatname"
 dgvRes.TextMatrix(0, 7) = "gesendet"
 
 
 Exit Sub
LOKAL_ERROR:
     

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "RettungHistorie"
    Fehler.gsFunktion = "ClearGridView"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub



Function UPDATE_STEUERKI() As Boolean
On Error GoTo LOKAL_ERROR

   UPDATE_STEUERKI = False
    
   Dim dtnam As String
   dtnam = dgvRes.TextMatrix(dgvRes.Row, 2)
    
   gdBase.Execute ("UPDATE STEUERKI SET gesendet = true WHERE DATNAME='" & dtnam & "'")
     
     
   UPDATE_STEUERKI = True
   
Exit Function

LOKAL_ERROR:

    UPDATE_STEUERKI = False
       
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "RettungHistorie"
    Fehler.gsFunktion = "UPDATE_STEUERKI"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function

 

 

Private Sub optDatum_Click()

 DTPicker1.Enabled = True
 txtDname.Text = ""
 txtDname.Enabled = False
 optName.value = False
 
End Sub

Private Sub optName_Click()
 
 DTPicker1.Enabled = False
 txtDname.Enabled = True
 optDatum.value = False
 
End Sub

Private Sub txtDname_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
 
   If Len(txtDname.Text) <> 0 Then
   
     Dim tmpVal As String
      
     tmpVal = Trim(txtDname.Text)
     txtDname.Text = ""
     
     tmpVal = Replace(tmpVal, ".", "")
     tmpVal = Replace(tmpVal, ",", "")
     tmpVal = Replace(tmpVal, " ", "")
     
     If Not Left(tmpVal, 1) = "F" And Not Left(tmpVal, 1) = "f" Then
       Exit Sub
     End If
     
     If Len(tmpVal) > 1 Then
     
        Dim nachF As String
        nachF = Mid(tmpVal, 2, Len(tmpVal) - 1)
     
        If Not IsNumeric(nachF) Then
           Exit Sub
        End If
        
     End If
     
     
     If Len(tmpVal) > 8 Then
       tmpVal = Mid(tmpVal, 1, 8)
       txtDname.Text = tmpVal
     Else
       txtDname.Text = tmpVal
     End If
     
    txtDname.Text = UCase(txtDname.Text)
    txtDname.SelStart = Len(txtDname.Text)
    txtDname.SetFocus
   
   End If
      
      
      
       
Exit Sub

LOKAL_ERROR:
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "RettungHistorie"
    Fehler.gsFunktion = "txtDname_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

 
