VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FDateienRettung 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "FDateienRettung"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Historie"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtFDatName1 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton btnRetten 
      Caption         =   "F-Datei wiederherstellen"
      Enabled         =   0   'False
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
      Left            =   2640
      TabIndex        =   11
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox txtPasswort 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   3120
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   113901569
      CurrentDate     =   44482
   End
   Begin VB.TextBox txtFDatName2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "* * * * *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   6615
   End
   Begin VB.Label Label9 
      Caption         =   ".lzh"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "( Format: F0000000.lzh )"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Passwort :"
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
      Left            =   1080
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblKasse 
      Caption         =   "*****"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Kasse :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Datum :"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "F-DateiName :"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblFil 
      Caption         =   "*****"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Filiale :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "F-Datei Wiederherstellung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "FDateienRettung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DateiZumSenden As String

Private Sub btnRetten_Click()
 On Error GoTo LOKAL_ERROR
  
  If Not DateValue(DTPicker1.value) < DateValue(Now) Then
       
       'Hinweis: dieses Formular ist zur Rettung der(vorherigen) F-Dateien, deswegen muss man das Datum von heute
       '         nicht auswählen, weil der Tagesabschluss für heute vielleicht noch nicht durchgeführt wurde.
       MsgBox ("bitte einen vorherigen Tag auswählen !")
       Exit Sub
    
  End If
 
 'dieses Formular ist mit einem Passwort geschützt, weil nur wir (KISS Mitarbeiter) die FDateien-Rettung bei den Kunden durchführen dürfen.
 If txtPasswort.Text = "fkiss2021" Then
   
  Dim ress As Integer
  ress = MsgBox("möchten Sie starten ?", vbQuestion + vbYesNo, "FDatei Rettung")
                       
    If ress = vbYes Then
     
 
     '''''''''''''''''''''''''''''''F-Dateien Rettung <<<<<<<<<<<<<< START
      
      txtFDatName2.Enabled = False
      DTPicker1.Enabled = False
      txtPasswort.Enabled = False
      btnRetten.Enabled = False
      
     'Tabellen in [ gcDBPfad\FDateien Rettung\FZ.mdb ]  füllen
      If Not TabellenMitDenDatenFuellen Then
       Exit Sub
      End If
      
      'in diesem Directory wird die Ausgabe der RettungsProzess am Ende zur Verfügung gestellt.
      If Dir(gcDBPfad & "\END Rettung\" & CStr(DTPicker1.value), vbDirectory) = "" Then
         MkDir gcDBPfad & "\END Rettung\" & CStr(DTPicker1.value)
      End If
      
     lblStatus.Caption = "F-Datei wird gezippt . . ."
     lblStatus.Refresh
     
     Text1.Text = "0"
     Zip_Folder "", gcDBPfad & "\FDateien Rettung", gcDBPfad & "\END Rettung\" & CStr(DTPicker1.value) & "\" & txtFDatName1.Text & txtFDatName2.Text & ".lzh", Text1
     
     lblStatus.Caption = "Rettung Fertig"
     lblStatus.Refresh
     
    '''''''''''''''''''''''''''''''F-Dateien Rettung <<<<<<<<<<<<<<<<< ENDE
    
    
    
    DateiZumSenden = txtFDatName1.Text & txtFDatName2.Text & "_Rettung_KASSE_" & gcKasNum & "_" & CStr(DTPicker1.value) & ".lzh"
    
    
    
    
    
    'gerettete F-Datei an die Zentrale absenden <<<<<<<<<<<<<<<<<<<<<<< START
     
         Dim ireslt1 As Integer
         ireslt1 = MsgBox("möchten Sie die gerettete F-Datei an die Zentrale absenden ?", vbQuestion + vbYesNo, "gerettete FDatei an die Zentrale")
                       
         If ireslt1 = vbYes Then
         
                'in diesem Directory steht die gerettete F-Datei, die an die Zentrale geschickt wird.
                If Dir(gcDBPfad & "\RettungAnZentrale", vbDirectory) = "" Then
                       MkDir gcDBPfad & "\RettungAnZentrale"
                Else
                       If Not Dir(gcDBPfad & "\RettungAnZentrale\*.*") = "" Then
                        Kill gcDBPfad & "\RettungAnZentrale\*.*"
                       End If
                      
                End If
        
        
                Dim lRet  As Long
                Dim lfail As Long
                 
                
                lRet = CopyFile(gcDBPfad & "\END Rettung\" & CStr(DTPicker1.value) & "\" & txtFDatName1.Text & txtFDatName2.Text & ".lzh", gcDBPfad & "\RettungAnZentrale\" & txtFDatName1.Text & txtFDatName2.Text & ".lzh", lfail)
                    
                If Not lRet = 1 Then
                     
                     MsgBox ("Kopieren fehlgeschlagen !!! " & vbNewLine & vbNewLine & "Dateiname: " & txtFDatName1.Text & txtFDatName2.Text & ".lzh" & vbNewLine & vbNewLine & " von:" & vbNewLine & gcDBPfad & "\END Rettung\" & CStr(DTPicker1.value) & vbNewLine & vbNewLine & "in: " & vbNewLine & gcDBPfad & "\RettungAnZentrale")
                     Exit Sub
                     
                Else
                
gerettete_FDatei_An_Die_Zentrale:

                       geretteteF_DateiErfolgreichAbgeschickt = False
                       giKissFtpMode = 50
                       frmWKL38.Show 1
                       
                       
                       If geretteteF_DateiErfolgreichAbgeschickt Then
                       
                            lblStatus.Caption = "erfolgreich an die Zentrale abgeschickt."
                            lblStatus.Refresh
                            lblStatus.ForeColor = vbGreen
                            
                            
                            'in der Tabelle STEUERKI schreiben
                            If Not InSTEUERKI_schreiben(True) Then
                             MsgBox ("das Schreiben in der Tabelle ' STEUERKI ' war fehlgeschlagen !!!")
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
                                    Exit Sub
                             End If
                              
                       End If
                       
                       
                End If
        
         Else
          
                'in der Tabelle STEUERKI schreiben
                 If Not InSTEUERKI_schreiben(False) Then
                  MsgBox ("das Schreiben in der Tabelle ' STEUERKI ' war fehlgeschlagen !!!")
                 End If
                  
         End If
     
     'gerettete F-Datei an die Zentrale absenden <<<<<<<<<<<<<<<<<<<<<<< ENDE
     
    
    
    
    Else
     Exit Sub
    End If
          
    
      
 Else
  MsgBox ("falsches Passwort !!!")
  Exit Sub
 End If
 


Exit Sub
LOKAL_ERROR:
     
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "btnRetten_Click"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub



Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR
 
'
'   If DateValue(DTPicker1.value) < DateValue(Now) Then
'     MsgBox ("weniger")
'   Else
'     MsgBox ("größer")
'   End If
 
 
Exit Sub

LOKAL_ERROR:
    

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub

Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR

If txtPasswort.Text = "fkiss2021" Then
    
     'If GibtEsSchonGerettetenDateien Then
                'Historie der geretteten FDateien
                 
                 RettungHistorie.Left = (frmWKL00.ScaleWidth - RettungHistorie.Width) / 2
                 RettungHistorie.Top = (frmWKL00.ScaleHeight - RettungHistorie.Height) / 2
                 RettungHistorie.Show 1
                 
     'Else
            
         'MsgBox ("keine geretteten F-Dateien gefunden (Historie ist leer) !!!")
            
     'End If
 
Else

    MsgBox ("falsches Passwort !!!")

End If


Exit Sub
LOKAL_ERROR:
     
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
     
    btnRetten.Enabled = False
     
   'Sperr Datei ' EineF-DateiWirdGerettet.txt ' erstellen ( wenn diese Datei existiert, wird die überschrieben ).
    SperrMichBitte
    
    If Not STEUERKI_Erweitern Then
     Exit Sub
    End If
    
   
    
    If Len(CStr(gcFilNr)) = 2 Then
        txtFDatName1.Text = "F" & CStr(gcFilNr)
    Else
        txtFDatName1.Text = "F0" & CStr(gcFilNr)
    End If
    
    lblFil.Caption = gcFilNr
    lblFil.Refresh
    
    lblKasse.Caption = gcKasNum
    lblFil.Refresh
    
    
    If Dir(gcDBPfad & "\FDateien Rettung", vbDirectory) = "" Then
    
             'Directory existiert nicht ... erstelle es (in diesem Directory wird die F-Dateien geretett)
             MkDir gcDBPfad & "\FDateien Rettung"
    Else
            'wenn Directory 'FDateien Rettung' nicht leer ist ... dann entleeren
             If Not Dir(gcDBPfad & "\FDateien Rettung\*.*") = "" Then
              Kill gcDBPfad & "\FDateien Rettung\*.*"
             End If
            
    End If
    
    
    If Dir(gcDBPfad & "\END Rettung", vbDirectory) = "" Then
            'Directory existiert nicht ... erstelle es (in diesem Directory wird die geretetten F-Dateien gespeichert )
            MkDir gcDBPfad & "\END Rettung"
    End If
    
    
    'prüf mal, ob FZ.mdb existiert ( FZ.mdb ist die Datenbank, die die leere Default-Tabellen enthält(sind die Tabellen nicht leer, dann entleere die unten in dieser Methode). diese Tabellen (nicht alle) würden am Ende dieses Prozess gefüllt werden. und dann wird die ganze FZ.mdb als F-Datei im Ordner (FDateien Rettung) gezippt)
    If Not FileExists(gcDBPfad & "\FZ.mdb") Then
    
     MsgBox ("die Datenbank FZ.mdb, die leere Tabellen enthält, muss in diesem Pfad existieren." & vbNewLine & vbNewLine & gcDBPfad)
     Exit Sub
     
    End If
    
    
    'prüf mal, ob FZ.mdb erreichbar ist
    Dim FZtest As Database
    Set FZtest = OpenDatabase(gcDBPfad & "\FZ.mdb", False, False, "MS Access;PWD=XYC6T349G6")
    FZtest.Close 'erreichbar
    
    
    'jetzt die Tabellen in FZ.mdb entleeren (jetzt ist der Aufruf dieser Funktion unnötig aber vielleicht nachher wäre nötig)
    'If dieTabellenInDerDatenbankFZentleeren Then
    
    'End If
    
    
    
    
    ' kopiere FZ.mdb im Ordner ' FDateien Rettung '   <<<<<<< START
        Dim lRet  As Long
        Dim lfail As Long
            
        lRet = CopyFile(gcDBPfad & "\FZ.mdb", gcDBPfad & "\FDateien Rettung\FZ.mdb", lfail)
            
        If Not lRet = 1 Then
                MsgBox ("Kopieren fehlgeschlagen !" & vbNewLine & vbNewLine & "von: " & gcDBPfad & "\FZ.mdb" & "in: " & gcDBPfad & "\FDateien Rettung\FZ.mdb")
                Exit Sub
        Else
               
        End If
    
    ' kopiere FZ.mdb im Ordner ' FDateien Rettung '   <<<<<<< ENDE
      
      
      
      
      
      'naechste LFNR generieren  <<<<<<< START
        Dim slf As String
        Dim lLFNR As Long
         
        lLFNR = lfnrErmitteln("F")
        lLFNR = lLFNR + 1
      
        slf = CStr(lLFNR)
        slf = Space(5 - Len(slf)) & slf
        slf = SwapStr(slf, " ", "0")
        
        txtFDatName2.Text = slf
      'naechste LFNR generieren  <<<<<<< ENDE
      
  
  

Exit Sub
LOKAL_ERROR:
    
    btnRetten.Enabled = False

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub



Function dieTabellenInDerDatenbankFZentleeren() As Boolean
On Error GoTo LOKAL_ERROR

   dieTabellenInDerDatenbankFZentleeren = False

   Dim i As Integer
   Dim lcount As Integer
    
   Dim tabname As String
   
   Dim dbFZ As Database
   Set dbFZ = OpenDatabase(gcDBPfad & "\FZ.mdb", False, False, "MS Access;PWD=XYC6T349G6")
 

   dbFZ.TableDefs.Refresh
   i = dbFZ.TableDefs.Count
   
   ''''''' durch alle Tabellen schleifen  <<<<<<<<<< START
   For lcount = 0 To i - 1
       
       tabname = dbFZ.TableDefs(lcount).name
       If Left(tabname, 4) <> "MSys" And tabname <> "f_9E8203D96A754B0890DAF9414007C362_Data" Then
         dbFZ.Execute "DELETE FROM " & tabname, dbFailOnError
       End If
       
   Next lcount
  '''''' durch alle Tabellen schleifen   <<<<<<<<<<< ENDE
  dbFZ.Close
  
  dieTabellenInDerDatenbankFZentleeren = True
  

Exit Function

LOKAL_ERROR:

    dieTabellenInDerDatenbankFZentleeren = False
     

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "dieTabellenInDerDatenbankFZentleeren"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function

 

 
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
  
  If FileExists(gcDBPfad & "\EineF-DateiWirdGerettet.txt") Then
      Kill gcDBPfad & "\EineF-DateiWirdGerettet.txt"
  End If
 
Exit Sub

LOKAL_ERROR:
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

 

Private Sub txtFDatName2_Change()
On Error GoTo LOKAL_ERROR
   
  
 Dim textval As String
  
 textval = Trim(txtFDatName2.Text)
 textval = Replace(textval, ".", "")
 textval = Replace(textval, ",", "")
 
  If IsNumeric(textval) And Len(textval) Then
      txtFDatName2.Text = CStr(textval)
    Else
      txtFDatName2.Text = ""
  End If

 If Len(txtFDatName2.Text) = 5 Then
      btnRetten.Enabled = True
     Else
      btnRetten.Enabled = False
 End If
 
 
 
Exit Sub

LOKAL_ERROR:
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "txtFDatName2_Change"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
 
 

Function SperrMichBitte() As Boolean
On Error GoTo LOKAL_ERROR

  SperrMichBitte = False
 
  Dim iFileNo As Integer
  iFileNo = FreeFile
  Open gcDBPfad & "\EineF-DateiWirdGerettet.txt" For Output As #iFileNo
  Close #iFileNo
                
  SperrMichBitte = True
  

Exit Function

LOKAL_ERROR:

    SperrMichBitte = False
     

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "SperrMichBitte"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function



Function TabellenMitDenDatenFuellen() As Boolean
On Error GoTo LOKAL_ERROR

   TabellenMitDenDatenFuellen = False
   
   lblStatus.ForeColor = vbBlack
   
   Dim ZiehlDatenbank As String
   ZiehlDatenbank = gcDBPfad & "\FDateien Rettung\FZ.mdb"
   
   ''''''''''''''''''''''''''''''''''''''''''''''' Haupt Tabellen ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   '1) KKZAHLO,KKZAHLTEO
   lblStatus.Caption = "Tabelle [ KKZAHL ] wird abgefragt . . ."
   lblStatus.Refresh
   
   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KKZAHLO SELECT * FROM KKZAHL WHERE Datevalue(ADATE)=CDate('" & DTPicker1.value & "') AND KASNUM=" & gcKasNum & " AND FILIALE=" & gcFilNr)
   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KKZAHLTEO SELECT * FROM KKZAHL WHERE Datevalue(ADATE)=CDate('" & DTPicker1.value & "') AND KASNUM=" & gcKasNum & " AND FILIALE=" & gcFilNr)
   gdBase.Execute ("UPDATE [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KKZAHLO SET SENDOK=false")
   gdBase.Execute ("UPDATE [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KKZAHLTEO SET SENDOK=false")
    
   '2) LASTZAHLO
   lblStatus.Caption = "Tabelle [ LASTZAHL ] wird abgefragt . . ."
   lblStatus.Refresh
   
   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].LASTZAHLO SELECT * FROM LASTZAHL WHERE Datevalue(ADATE)=CDate('" & DTPicker1.value & "') AND KASNUM=" & gcKasNum & " AND FILIALE=" & gcFilNr)
   gdBase.Execute ("UPDATE [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].LASTZAHLO SET SENDOK=false")
    
   '3) ABSCHOPFO
   lblStatus.Caption = "Tabelle [ ABSCHOPF ] wird abgefragt . . ."
   lblStatus.Refresh
   
   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].ABSCHOPFO SELECT * FROM ABSCHOPF WHERE Datevalue(ADATE)=CDate('" & DTPicker1.value & "') AND KASNUM=" & gcKasNum & " AND FILIALE=" & gcFilNr)
   gdBase.Execute ("UPDATE [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].ABSCHOPFO SET SENDOK=false")
    
   '4) DUKATENBO
   lblStatus.Caption = "Tabelle [ DUKATENB ] wird abgefragt . . ."
   lblStatus.Refresh
   
   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].DUKATENBO SELECT * FROM DUKATENB WHERE Datevalue(ADATE)=CDate('" & DTPicker1.value & "') AND KASNUM=" & gcKasNum & " AND FILIALE=" & gcFilNr)
   gdBase.Execute ("UPDATE [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].DUKATENBO SET SENDOK=false")
    
   '5) AFCBUCH
   lblStatus.Caption = "Tabelle [ KASSJOUR ] wird abgefragt . . ."
   lblStatus.Refresh
   
   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].AFCBUCH SELECT '0' as AFLAG,MENGE as AMENGE,PREIS as APREIS,ARTNR as AARTNR,BEZEICH as ABEZEICH,ADATE,AZEIT,'0' as AMWST,KUNDNR as AKUNUM,MWST as AMWSK,BELEGNR,KASNUM,'false' as KREDITFLAG,'0' as BUCHFLAG,PREIS as AALTPREIS,VKPR as AVKPR,EKPR as ALEKPR,LINR,KK_ART,BEST1 as BESTAND,'0' as ZHLGGUTSCH,UMS_OK,FILIALE as FILIALNR,'' as BONUS_OK,'A' as SYNSTATUS,BEDIENER as aBEDNU FROM KASSJOUR WHERE Datevalue(ADATE)=CDate('" & DTPicker1.value & "') AND KASNUM=" & gcKasNum & " AND FILIALE=" & gcFilNr)
   gdBase.Execute ("UPDATE [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].AFCBUCH AF INNER JOIN ARTIKEL AR ON AF.AARTNR=AR.ARTNR SET AF.BONUS_OK=AR.BONUS_OK")
       
       
   ''''''''''''''''''''''''''''''''''''''''''''''' andere Tabellen '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ''''''''' für die folgende weitere Tabellen musst du WHERE Datum= , KASNUM= ,FILIALE =  hinzufügen (wenn die Spalten Datum,KASNUM,FILIALE in diesen Tabellen existieren)
   
   
   ''6)KUN_OUT
'   lblStatus.Caption = "6.Tabelle [ KUNDEN ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KUN_OUT SELECT * from KUNDEN where SYNSTATUS = 'A' or  SYNSTATUS = 'E' or  SYNSTATUS = 'D' ")
'   gdBase.Execute ("Update KUNDEN set STATUS = 'N' where STATUS <> 'N'")
'   gdBase.Execute ("Update KUNDEN set SYNSTATUS = 'N' where SYNSTATUS <> 'N'")
   
   
   ''7)BED_OUT
'   lblStatus.Caption = "7.Tabelle [ Bedname ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].BED_OUT Select * from Bedname where SYNSTATUS = 'A' or  SYNSTATUS = 'E' or  SYNSTATUS = 'D' ")
   
   
   ''8)GUT_OUT
'   lblStatus.Caption = "8.Tabelle [ Gutsch ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   If Not gbKL_LIVEGUTSCHEIN Then
'      gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].GUT_OUT Select * from Gutsch ")
'   End If
    
   
   ''9)LogY
'   lblStatus.Caption = "9.Tabelle [ Steuerki ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].LogY SELECT top 10 Datname,datum, " & CByte(gcFilNr) & "  as filiale from Steuerki where Left(datname,1) = 'Y' order by lfnr desc")
   
   
   ''10)FILTAU
'   lblStatus.Caption = "10.Tabelle [ TAUSCH ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].FILTAU Select * from TAUSCH where SENDOK = False")
'   gdBase.Execute ("update TAUSCH set SENDOK = True where SENDOK = False")
   
   ''11)EANPOUT
'   lblStatus.Caption = "11.Tabelle [ EANPROT ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].EANPOUT Select * from EANPROT where SENDOK = False")
'   gdBase.Execute ("update EANPROT set SENDOK = True where SENDOK = False")
   
   ''12)KVKPR1POUT
'   lblStatus.Caption = "12.Tabelle [ KVKPR1PROT ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KVKPR1POUT Select * from KVKPR1PROT where SENDOK = False")
'   gdBase.Execute ("update KVKPR1PROT set SENDOK = True where SENDOK = False")
   
   ''13)BESTPOUT
'   lblStatus.Caption = "13.Tabelle [ BESTPROT ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].BESTPOUT Select * from BESTPROT where SENDOK = False")
'   gdBase.Execute ("update BESTPROT set SENDOK = True where SENDOK = False")
     
   ''14)RETOUT    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <--------------------------------------------------------------------------------
'   lblStatus.Caption = "14.Tabelle [ Retoure ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].RETOUT Select * from Retoure where SENDOK = False")
'   gdBase.Execute ("update Retoure set SENDOK = True where SENDOK = False")
   
   ''15)KBOUT
'   lblStatus.Caption = "15.Tabelle [ KUNDBEST ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KBOUT Select * from KUNDBEST where SENDOK = False")
'   gdBase.Execute ("update KUNDBEST set SENDOK = True where SENDOK = False")
   
   ''16)GUTZO    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <---------------------------------------------------------------------------------
'   lblStatus.Caption = "16.Tabelle [ GUTZ ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].GUTZO Select * from GUTZ where SENDOK = False")
'   gdBase.Execute ("update GUTZ set SENDOK = True where SENDOK = False")
   
   ''17)GUHISO    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <--------------------------------------------------------------------------------
'   lblStatus.Caption = "17.Tabelle [ GUHIS ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].GUHISO Select * from GUHIS where SENDOK = False")
'   gdBase.Execute ("update GUHIS set SENDOK = True where SENDOK = False")
   
   ''18)KUNDEDELO
'   lblStatus.Caption = "18.Tabelle [ KUNDEDEL ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KUNDEDELO Select * from KUNDEDEL where SENDOK = False")
'   gdBase.Execute ("update KUNDEDEL set SENDOK = True where SENDOK = False")
   
   ''19)ZUGANGFO
'   lblStatus.Caption = "19.Tabelle [ ZUGANGF ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].ZUGANGFO Select * from ZUGANGF where SENDOK = False")
'   gdBase.Execute ("update ZUGANGF set SENDOK = True where SENDOK = False")
   
   ''20)LASTZAHLTEO    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <---------------------------------------------------------------------------
'   lblStatus.Caption = "20.Tabelle [ LASTZAHLTE ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].LASTZAHLTEO Select * from LASTZAHLTE where SENDOK = False")
'   gdBase.Execute ("update LASTZAHLTE set SENDOK = True where SENDOK = False")
   
   ''21)MBORDERO
'   lblStatus.Caption = "21.Tabelle [ MBORDER ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].MBORDERO Select * from MBORDER where SENDOK = False")
'   gdBase.Execute ("update MBORDER set SENDOK = True where SENDOK = False")
   
   ''22)MBORDERDELO
'   lblStatus.Caption = "22.Tabelle [ MBORDERDEL ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'
'
'     If NewTableSuchenDBKombi("MBORDERDEL", gdBase) = True Then
'
'
'       Dim fzDb As Database
'       Set fzDb = OpenDatabase(ZiehlDatenbank, False, False, "MS Access;PWD=XYC6T349G6")
'
'       If NewTableSuchenDBKombi("MBORDERDELO", fzDb) = True Then
'
'          gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].MBORDERDELO Select * from MBORDERDEL where SENDOK = False")
'          gdBase.Execute ("update MBORDERDEL set SENDOK = True where SENDOK = False")
'
'       Else
'
'          gdBase.Execute ("SELECT * INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].MBORDERDELO from MBORDERDEL where SENDOK = False")
'          gdBase.Execute ("update MBORDERDEL set SENDOK = True where SENDOK = False")
'
'       End If
'
'       fzDb.Close
'
'
'     End If
    
   
    
     
  
   
   ''23)GEMZO    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <--------------------------------------------------------------------------------
'   lblStatus.Caption = "23.Tabelle [ GEMZ ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].GEMZO Select * from GEMZ where SENDOK = False")
'   gdBase.Execute ("update GEMZ set SENDOK = True where SENDOK = False")
   
   ''24)MAILFBO
'   lblStatus.Caption = "24.Tabelle [ MAILFB ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].MAILFBO Select * from MAILFB where SENDOK = False")
'   gdBase.Execute ("update MAILFB set SENDOK = True where SENDOK = False")
   
   ''25)BONUSNRO
'   lblStatus.Caption = "25.Tabelle [ BONUSNR ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   If NewTableSuchenDBKombi("BONUSNR", gdBase) = True Then
'    gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].BONUSNRO Select * from BONUSNR where SENDOK = False")
'    gdBase.Execute ("update BONUSNR set SENDOK = True where SENDOK = False")
'   End If
   
   ''26)GANALYSEALLO
'   lblStatus.Caption = "26.Tabelle [ GANALYSEALL ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].GANALYSEALLO Select * from GANALYSEALL where SENDOK = False")
'   gdBase.Execute ("update GANALYSEALL set SENDOK = True where SENDOK = False")
   
   ''27)KABUCHO    [diese Tabelle enthält die Spalten Datum,FILIALE,KASNUM] <--------------------------------------------------------------------------------
'   lblStatus.Caption = "27.Tabelle [ KABUCH ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KABUCHO Select * from KABUCH where SENDOK = False")
'   gdBase.Execute ("update KABUCH set SENDOK = True where SENDOK = False")
   
   ''28)BONUS_SYSO
'   lblStatus.Caption = "28.Tabelle [ BONUS_SYS ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].BONUS_SYSO Select * from BONUS_SYS where SENDOK = False")
'   gdBase.Execute ("update BONUS_SYS set SENDOK = True where SENDOK = False")
    
   ''29)UNTERWFO
'   lblStatus.Caption = "29.Tabelle [ UNTERWF ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].UNTERWFO Select * from UNTERWF where SENDOK = False")
'   gdBase.Execute ("update UNTERWF set SENDOK = True where SENDOK = False")
   
   ''30)ALTERGO    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <--------------------------------------------------------------------------------
'   lblStatus.Caption = "30.Tabelle [ ALTERG ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].ALTERGO Select * from ALTERG where SENDOK = False")
'   gdBase.Execute ("update ALTERG set SENDOK = True where SENDOK = False")
   
   ''31)NEINVKO
'   lblStatus.Caption = "31.Tabelle [ NEINVK ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].NEINVKO Select * from NEINVK where SENDOK = False")
'   gdBase.Execute ("update NEINVK set SENDOK = True where SENDOK = False")
   
   ''32)KREDITZAO
'   lblStatus.Caption = "32.Tabelle [ KREDITZA ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KREDITZAO Select * from KREDITZA where SENDOK = False")
'   gdBase.Execute ("update KREDITZA set SENDOK = True where SENDOK = False")
   
   ''33)FEEDBO
'   lblStatus.Caption = "33.Tabelle [ FEEDB ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].FEEDBO Select * from FEEDB where SENDOK = False")
'   gdBase.Execute ("update FEEDB set SENDOK = True where SENDOK = False")
   
   ''34)FEEDB_TRANSO
'   lblStatus.Caption = "34.Tabelle [ FEEDB_TRANS ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].FEEDB_TRANSO Select * from FEEDB_TRANS where SENDOK = False")
'   gdBase.Execute ("update FEEDB_TRANS set SENDOK = True where SENDOK = False")
   
   
   ''35)FEEDBFO
'   lblStatus.Caption = "35.Tabelle [ FEEDBF ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].FEEDBFO Select * from FEEDBF where SENDOK = False")
'   gdBase.Execute ("update FEEDBF set SENDOK = True where SENDOK = False")
   
   ''36)KAEINAUSFO    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <-----------------------------------------------------------------------
'   lblStatus.Caption = "36.Tabelle [ KAEINAUSF ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KAEINAUSFO Select * from KAEINAUSF where SENDOK = False")
'   gdBase.Execute ("update KAEINAUSF set SENDOK = True where SENDOK = False")
   
   ''37)BESTOUT
'   lblStatus.Caption = "37.Tabelle [ BESTAEND ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].BESTOUT Select * from BESTAEND where SENDOK = False")
'   gdBase.Execute ("update BESTAEND set SENDOK = True where SENDOK = False")
   
   ''38)BARGOUT
'   lblStatus.Caption = "38.Tabelle [ BARGELD ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].BARGOUT Select * from BARGELD where SENDOK = False")
'   gdBase.Execute ("update BARGELD set SENDOK = True where SENDOK = False")
   
   ''39)STORNO2O    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <-------------------------------------------------------------------------
'   lblStatus.Caption = "39.Tabelle [ STORNO2 ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].STORNO2O Select * from STORNO2 where SENDOK = False")
'   gdBase.Execute ("update STORNO2 set SENDOK = True where SENDOK = False")
     
   ''40)ARTDETO    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <--------------------------------------------------------------------------
'   lblStatus.Caption = "40.Tabelle [ ARTDET ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].ARTDETO Select * from ARTDET where SENDOK = False")
'   gdBase.Execute ("update ARTDET set SENDOK = True where SENDOK = False")
     
   ''41)KASSBEDPO    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <------------------------------------------------------------------------
'   lblStatus.Caption = "41.Tabelle [ KASSBEDP ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KASSBEDPO Select * from KASSBEDP where SENDOK = False")
'   gdBase.Execute ("update KASSBEDP set SENDOK = True where SENDOK = False")
     
   ''42)AFCSTATPO    [diese Tabelle enthält die Spalten ADATE,KASNUM] <--------------------------------------------------------------------------------
'   lblStatus.Caption = "42.Tabelle [ AFCSTATP ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].AFCSTATPO Select * from AFCSTATP where SENDOK = False")
'   gdBase.Execute ("update AFCSTATP set SENDOK = True where SENDOK = False")
     
   ''43)DTAOUT
'   lblStatus.Caption = "43.Tabelle [ DTA ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   If gbECTOZ Then
'
'       Dim fzDb2 As Database
'       Set fzDb2 = OpenDatabase(ZiehlDatenbank, False, False, "MS Access;PWD=XYC6T349G6")
'
'       If NewTableSuchenDBKombi("DTAOUT", fzDb2) = True Then
'
'         gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].DTAOUT Select * from DTA where SENDOK = False")
'         gdBase.Execute ("update DTA set SENDOK = True where SENDOK = False")
'
'       Else
'
'          gdBase.Execute ("SELECT * INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].DTAOUT from DTA where SENDOK = False")
'          gdBase.Execute ("update DTA set SENDOK = True where SENDOK = False")
'
'       End If
'
'       fzDb.Close
'
'   End If
'
   ''44)STE_OUT
'   lblStatus.Caption = "44.Tabelle [ STEMPEL ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].STE_OUT Select * from STEMPEL where SENDOK = False")
'   gdBase.Execute ("update STEMPEL set SENDOK = True where SENDOK = False")
     
   ''45)KOLOUT    [diese Tabelle enthält die Spalten ADATE,FILIALE,KASNUM] <----------------------------------------------------------------------------------
'   lblStatus.Caption = "45.Tabelle [ KOLLVERK ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KOLOUT Select * from KOLLVERK where SENDOK = False")
'   gdBase.Execute ("UPDATE KOLLVERK SET SENDOK = True where SENDOK = False")
'
   ''46)KASSBONO    [diese Tabelle enthält die Spalten Datum,FILIALE,KASNUM] <--------------------------------------------------------------------------------
'   lblStatus.Caption = "46.Tabelle [ KASSBON ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KASSBONO Select * from KASSBON where SENDOK = False")
'   gdBase.Execute ("update KASSBON set SENDOK = True where SENDOK = False")
     
   ''47)KUNDENBONUSO
'   lblStatus.Caption = "47.Tabelle [ KUNDENBONUS ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].KUNDENBONUSO Select * from KUNDENBONUS where SENDOK = False")
'   gdBase.Execute ("update KUNDENBONUS set SENDOK = True where SENDOK = False")
     
   ''48)Z_OUT
'   lblStatus.Caption = "48.Tabelle [ ARTIKEL ] wird abgefragt . . ."
'   lblStatus.Refresh
'
'   gdBase.Execute ("INSERT INTO [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].Z_OUT Select ARTNR,BESTAND,MINBEST,KVKPR1 from ARTIKEL where BESTAND <> 0 or MINBEST > 0 or KVKPR1 > 0")
'   gdBase.Execute ("update [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].Z_out Set MINBEST = 0 where MINBEST is null")
'   gdBase.Execute ("update [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].Z_out Set BESTAND = 0 where BESTAND is null")
'   gdBase.Execute ("update [MS Access;PWD=XYC6T349G6;Database=" & ZiehlDatenbank & "].Z_out Set KVKPR1 = 0 where KVKPR1 is null")
     
    
     
     
   TabellenMitDenDatenFuellen = True
   
Exit Function

LOKAL_ERROR:

    TabellenMitDenDatenFuellen = False
     
    lblStatus.ForeColor = vbRed
    lblStatus.Caption = "fehlgeschlagen ! ! !"
    lblStatus.Refresh

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "TabellenMitDenDatenFuellen"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function


Function STEUERKI_Erweitern() As Boolean
On Error GoTo LOKAL_ERROR

    STEUERKI_Erweitern = False
    
    Dim cSQL As String
     
   
    If Not SpalteInTabellegefundenNEW("STEUERKI", "IsRettungFDatei", gdBase) Then
       cSQL = "Alter table STEUERKI add column IsRettungFDatei bit"
       gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("STEUERKI", "RettungZumTag", gdBase) Then
       cSQL = "Alter table STEUERKI add column RettungZumTag DATETIME"
       gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("STEUERKI", "KASNUM", gdBase) Then
       cSQL = "Alter table STEUERKI add column KASNUM NUMBER"
       gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("STEUERKI", "EndDatname", gdBase) Then
       cSQL = "Alter table STEUERKI add column EndDatname varchar(50)"
       gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("STEUERKI", "gesendet", gdBase) Then
       cSQL = "Alter table STEUERKI add column gesendet bit"
       gdBase.Execute cSQL, dbFailOnError
       gdBase.Execute ("UPDATE STEUERKI SET gesendet = true WHERE IsRettungFDatei = false")
    End If
     
    
     
    STEUERKI_Erweitern = True
   
Exit Function

LOKAL_ERROR:

    STEUERKI_Erweitern = False
       
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "STEUERKI_Erweitern"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function

Function InSTEUERKI_schreiben(SchonGesendet As Boolean) As Boolean
On Error GoTo LOKAL_ERROR

   InSTEUERKI_schreiben = False
    
   Dim nlfnr As String
   Dim nDatname As String
   Dim nDatum As String
   Dim nRettungZumTag As String
   Dim nkasnum As String
   
   nlfnr = Right(txtFDatName2.Text, 4)
   nDatname = txtFDatName1.Text & txtFDatName2.Text
   nDatum = Format(DateValue(Now), "DD.MM.YYYY") & " " & Format(TimeValue(Now), "HH:MM:SS")
   nRettungZumTag = CStr(DTPicker1.value)
   nkasnum = CStr(gcKasNum)
   
   If SchonGesendet Then
   
    gdBase.Execute ("INSERT INTO STEUERKI (LFNR,DATNAME,DATUM,IsRettungFDatei,RettungZumTag,KASNUM,EndDatname,gesendet) VALUES ('" & nlfnr & "','" & nDatname & "','" & nDatum & "',true,'" & nRettungZumTag & "','" & nkasnum & "','" & DateiZumSenden & "',true)")
  
   Else
   
    gdBase.Execute ("INSERT INTO STEUERKI (LFNR,DATNAME,DATUM,IsRettungFDatei,RettungZumTag,KASNUM,EndDatname,gesendet) VALUES ('" & nlfnr & "','" & nDatname & "','" & nDatum & "',true,'" & nRettungZumTag & "','" & nkasnum & "','" & DateiZumSenden & "',false)")
    
   End If
   
    
     
   InSTEUERKI_schreiben = True
   
Exit Function

LOKAL_ERROR:

    InSTEUERKI_schreiben = False
       
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "InSTEUERKI_schreiben"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function




Function GibtEsSchonGerettetenDateien() As Boolean
On Error GoTo LOKAL_ERROR

GibtEsSchonGerettetenDateien = False
    
Dim rsrsHis As Recordset
Set rsrsHis = gdBase.OpenRecordset("SELECT count(*) as res FROM STEUERKI WHERE IsRettungFDatei=true")
If Not rsrsHis.EOF Then
    
   'rsrsHis.MoveFirst
   If Not IsNull(rsrsHis!Res) Then
       
       If CInt(rsrsHis!Res) > 0 Then
         GibtEsSchonGerettetenDateien = True
       End If
       
   End If
    
    
End If
   
     
Exit Function

LOKAL_ERROR:

    GibtEsSchonGerettetenDateien = False
       
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FDateienRettung"
    Fehler.gsFunktion = "GibtEsSchonGerettetenDateien"
    Fehler.gsFehlertext = "Im Programmteil Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function




