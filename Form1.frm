VERSION 5.00
Begin VB.Form FTPwechselAbbruch 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Budni-EDEKA Umstellung rückgängig machen"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Starten"
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
      Left            =   4680
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblAnzeig 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   7455
   End
   Begin VB.Label Label6 
      Caption         =   "Passwort :"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Ihre alte Budni-Kundennummer  :"
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
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "________________________________________________________________"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "machen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   " rückgängig"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Budni-EDEKA Umstellung "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "FTPwechselAbbruch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  
 lblAnzeig.Caption = ""
 lblAnzeig.Refresh
 
 'erstmal Formular Validierung             <<<<< START
 
 If Text1.Text <> "" Then
 
   If Text2.Text <> "" Then
   
        If Text2.Text = "brnas2030" Then
                     
                      'von EDEKA auf BUDNI wechseln
               
                       Me.Enabled = False
                        
                       If Not FileExists(gcDBPfad & "\neuBudniArtikel.mdb") Then
                         MsgBox ("Datei nicht gefunden" & vbNewLine & gcDBPfad & "\neuBudniArtikel.mdb")
                         Me.Enabled = True
                         Exit Sub
                       End If
                       
                       
                       If Not FileExists(gcDBPfad & "\TabellenSichern.mdb") Then
                         MsgBox ("Datei nicht gefunden" & vbNewLine & gcDBPfad & "\TabellenSichern.mdb")
                         Me.Enabled = True
                         Exit Sub
                       End If
                       
                       
                       If Not FileExists(gcDBPfad & "\KissdataMitEDEKANummer.mdb") Then
                         'jetzige Kissdata.mdb sichern
                         Dim Ret As Long
                         Dim lfail As Long
                         
                         lblAnzeig.Caption = "Kissdata.mdb wird als KissdataMitEDEKANummer.mdb gesichert ..."
                         lblAnzeig.Refresh
        
        
                         Ret = CopyFile(gcDBPfad & "\Kissdata.mdb", gcDBPfad & "\KissdataMitEDEKANummer.mdb", lfail)
                         If Ret <> 1 Then
                          lblAnzeig.Caption = "Kissdata.mdb sichern         * fehlgeschlagen *"
                          lblAnzeig.Refresh
                          Me.Enabled = True
                          Exit Sub
                         End If
                           
                       End If
                       
                         
                       
                       If BetroffeneTabellenZurucksetzen Then
                       
                            'Drop alle Tabellen in TabellenSichern.mdb
                            lblAnzeig.Caption = "Drop alle Tabellen in TabellenSichern.mdb . . ."
                            lblAnzeig.Refresh
                            
                            Dim dbs As Database
                            Dim tdf As TableDef
                            Set dbs = OpenDatabase(gcDBPfad & "\TabellenSichern.mdb")
                            For Each tdf In dbs.TableDefs
                            
                                If tdf.name = "ARTIKEL" Then
                                    dbs.Execute ("DROP TABLE ARTIKEL")
                                End If
                                
                                If tdf.name = "ARTLIEF" Then
                                    dbs.Execute ("DROP TABLE ARTLIEF")
                                End If
                              
                            Next tdf
                            
                            dbs.Close
                                            
                            'Drop FTPumzugFertig Tabelle
                            lblAnzeig.Caption = "drop Tabelle FTPumzugFertig . . ."
                            lblAnzeig.Refresh
                            gdBase.Execute ("DROP TABLE FTPumzugFertig")
                            
                            'jetzt delete KissdataMitEDEKANummer.mdb (ist nicht mehr nötig)
                            Kill gcDBPfad & "\KissdataMitEDEKANummer.mdb"
                            
                            
                            lblAnzeig.Caption = "Fertig ( erfolgreich )"
                            lblAnzeig.Refresh
                            
                            gbBudniNeuesFtpVerfahren = False
                            Me.Enabled = True
                            Unload Me
                            
                            'WINKISS auf neu Starten
                            MsgBox ("WINKISS wird jetzt beendet" & vbNewLine & "bitte starten Sie es auf neu.")
                            End
                       Else
                          lblAnzeig.Caption = ""
                          lblAnzeig.Refresh
                          Me.Enabled = True
                       End If
          
        Else
         MsgBox ("falsches Passwort !!!")
        End If
    
   
   Else
   MsgBox ("Bitte Passwort eingeben !!!")
   End If
  
 Else
  MsgBox ("Bitte Budni-Kundennummer eingeben !!!")
 End If
'erstmal Formular Validierung             <<<<< ENDE

End Sub

Private Sub Text1_Change()
 
 Dim iCount As Integer
 Dim tmpChar As String
 
 If Text1.Text <> "" Then
    
    For iCount = 1 To Len(Text1.Text)
     tmpChar = Mid(Text1.Text, iCount, 1)
     If InStr("1234567890", tmpChar) = 0 Then
      Text1.Text = ""
     End If
    
    Next iCount
    
 End If
 
End Sub


Private Function BetroffeneTabellenZurucksetzen() As Boolean
On Error GoTo LOKAL_ERROR

 'in dieser Funktion werden die damals(damals = nach der Umstellung von Budni auf Edeka) betroffenen Tabellen zurücksetzen

 BetroffeneTabellenZurucksetzen = False
 lblAnzeig.Caption = ""
 lblAnzeig.Refresh
 
 Dim Blinr As String
 Dim rsrs As Recordset
 Set rsrs = gdBase.OpenRecordset("select LINR FROM LISRT WHERE FORMAT='EDIBHSG'")
 If Not rsrs.EOF Then
    If Not IsNull(rsrs!linr) Then
     Blinr = rsrs!linr
    End If
 End If
 
 
 
 Dim FrauBrnasDateiPfad As String
 FrauBrnasDateiPfad = gcDBPfad & "\neuBudniArtikel.mdb"
 
 
 '1.LISRT Tabelle
 lblAnzeig.Caption = "1.die betroffene Tabelle LISRT zurücksetzen . . ."
 lblAnzeig.Refresh
 gdBase.Execute ("UPDATE LISRT SET KUNDNR='" & Text1.Text & "',GLN='" & Text1.Text & "',FORMAT='EDIBUDNI' WHERE FORMAT='EDIBHSG'")
 
 
 '2.ARTIKEL Tabelle
 lblAnzeig.Caption = "2.die betroffene Tabelle ARTIKEL zurücksetzen . . ."
 lblAnzeig.Refresh
 gdBase.Execute ("UPDATE ARTIKEL A INNER JOIN [MS Access;Database=" & FrauBrnasDateiPfad & "].neuBudniArtikelNr NA ON A.EAN=CStr(NA.[GTIN-Code]) AND A.libesnr=CStr(NA.[EDK Artik]) SET A.libesnr=CStr(NA.[Artikel]) WHERE NA.[GTIN-Code] is not null AND NA.[EDK Artik] is not null AND A.LINR=" & Blinr)
 
 
 '3.ARTLIEF Tabelle
 lblAnzeig.Caption = "3.die betroffene Tabelle ARTLIEF zurücksetzen . . ."
 lblAnzeig.Refresh
 gdBase.Execute ("UPDATE ARTLIEF AL INNER JOIN [MS Access;Database=" & FrauBrnasDateiPfad & "].neuBudniArtikelNr NA ON AL.libesnr=CStr(NA.[EDK Artik]) SET AL.libesnr=CStr(NA.[Artikel]) WHERE NA.[EDK Artik] is not null AND AL.LINR=" & Blinr)
        
        
 BetroffeneTabellenZurucksetzen = True
  
Exit Function
LOKAL_ERROR:

    BetroffeneTabellenZurucksetzen = False
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FTPwechselAbbruch"
    Fehler.gsFunktion = "BetroffeneTabellenZurucksetzen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function



