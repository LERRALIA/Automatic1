VERSION 5.00
Begin VB.Form TseEinstellungen 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "TSE Einstellungen"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "ok"
      Height          =   285
      Left            =   4800
      TabIndex        =   28
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   4080
      TabIndex        =   27
      Top             =   5880
      Width           =   615
   End
   Begin VB.CheckBox chkInterntZeit 
      BackColor       =   &H00C0C000&
      Caption         =   "Zeit vom Internet abfragen"
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
      Left            =   3960
      TabIndex        =   25
      Top             =   5520
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "?"
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
      Left            =   6240
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox chkAltDru 
      BackColor       =   &H00C0C000&
      Caption         =   "alter Druckmodus"
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
      Left            =   6720
      TabIndex        =   23
      Top             =   360
      Width           =   1815
   End
   Begin VB.CheckBox chkQrCode 
      BackColor       =   &H00C0C000&
      Caption         =   "QR-Code"
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
      Left            =   2400
      TabIndex        =   22
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton btnDatenExport 
      Caption         =   "Daten Exportieren"
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
      Left            =   2400
      TabIndex        =   21
      Top             =   4920
      Width           =   4935
   End
   Begin VB.CheckBox chkTSE_Status 
      BackColor       =   &H00C0C000&
      Caption         =   "TSE aktivieren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Width           =   3615
   End
   Begin VB.ComboBox comClients 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "TseEinstellungen.frx":0000
      Left            =   5520
      List            =   "TseEinstellungen.frx":0002
      TabIndex        =   19
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton btnSpeicher 
      Caption         =   "speichern"
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
      Left            =   5880
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton btnVerbinden 
      Caption         =   "Verbindung herstellen"
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
      Left            =   2400
      TabIndex        =   17
      Top             =   2640
      Width           =   4935
   End
   Begin VB.CommandButton btnTSEInfo 
      Caption         =   "TSE Info"
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
      Left            =   4440
      TabIndex        =   15
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton btnNeueClient 
      Caption         =   "neue Client"
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
      Left            =   2400
      TabIndex        =   14
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtTSE_SN 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Top             =   3840
      Width           =   4935
   End
   Begin VB.TextBox txtTSE_DeviceID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtTSE_TimeAdminpin 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5760
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtTSE_Adminpin 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtTSE_Port 
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtTSE_IP 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C000&
      Caption         =   "Time-out (1-120 s) : "
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
      Left            =   2400
      TabIndex        =   26
      Top             =   5930
      Width           =   1695
   End
   Begin VB.Label lblZeig 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1455
      Left            =   600
      TabIndex        =   16
      Top             =   6600
      Width           =   7695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      Caption         =   "S/N :"
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
      Left            =   1200
      TabIndex        =   12
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      Caption         =   "ClientID :"
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
      Left            =   4440
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      Caption         =   "DeviceID :"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Caption         =   "TimeAdminPin :"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "AdminPin :"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "TSE Einstellungen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Port :"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "IP :"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "TseEinstellungen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnDatenExport_Click()
    On Error GoTo LOKAL_ERROR
    
  If TSE_OK Then
    
    TSEDataExport.Left = (Me.Left + 600)
    TSEDataExport.Top = (Me.Top + 1000)
    TSEDataExport.Show 1
    
   Else
    MsgBox (TSE_Err)
  End If

 Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "btnDatenExport_Click"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub btnNeueClient_Click()
On Error GoTo LOKAL_ERROR
  Screen.MousePointer = 11
  
    Dim ans As String
    ans = InputBox("Kasse ID eingeben :", "neue Client regestrieren", 0)

    If Trim(ans) = "" Then
       
    Else
      NeuClientRegestrieren ans
      getAllClients
    End If
  Screen.MousePointer = 0
  
   Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "btnNeueClient_Click"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub btnSpeicher_Click()
    On Error GoTo LOKAL_ERROR
    
    lblZeig.ForeColor = vbBlack
    lblZeig.Caption = ""
        
    txtTSE_IP.BackColor = vbWhite
    txtTSE_Port.BackColor = vbWhite
    txtTSE_Adminpin.BackColor = vbWhite
    txtTSE_TimeAdminpin.BackColor = vbWhite
    comClients.BackColor = vbWhite
    
    
    If Trim(txtTSE_IP) = "" Then
        
        txtTSE_IP.BackColor = vbRed
        txtTSE_IP.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTSE_Port) = "" Then
        
        txtTSE_Port.BackColor = vbRed
        txtTSE_Port.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTSE_Adminpin) = "" Then
        
        txtTSE_Adminpin.BackColor = vbRed
        txtTSE_Adminpin.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTSE_TimeAdminpin) = "" Then
        
        txtTSE_TimeAdminpin.BackColor = vbRed
        txtTSE_TimeAdminpin.SetFocus
        Exit Sub
    End If
    
    If Trim(comClients.Text) = "" Then
        
        comClients.BackColor = vbRed
        comClients.SetFocus
        Exit Sub
    End If
    
  TSE_Einstellungen_Aktualisieren txtTSE_IP, txtTSE_Port, txtTSE_Adminpin, txtTSE_TimeAdminpin, comClients.Text, txtTSE_DeviceID.Text, txtTSE_SN.Text
 
 Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "btnSpeicher_Click"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub btnTSEInfo_Click()

 If TSE_OK Then
    TSEStorageInfo.Show 1
   Else
    MsgBox (TSE_Err)
 End If
   
End Sub
 
  
Private Sub chkAltDru_Click()
 On Error GoTo LOKAL_ERROR
 
 If chkAltDru.value = vbChecked Then
        
        altDruckModus = True
        SqlCmd = "UPDATE TSESettings SET AltDModus='1'"
        gdApp.Execute SqlCmd, dbFailOnError
        
    Else
    
        altDruckModus = False
        SqlCmd = "UPDATE TSESettings SET AltDModus='0'"
        gdApp.Execute SqlCmd, dbFailOnError
         
 End If
 
Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "chkAltDru_Click"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub chkInterntZeit_Click()
On Error GoTo LOKAL_ERROR

    If chkInterntZeit.value = vbChecked Then
        TSE_InternetZeitAbfragen = True
        SqlCmd = "UPDATE TSESettings SET ZeitVomInternet='1'"
        gdApp.Execute SqlCmd, dbFailOnError
        
    Else
    
        TSE_InternetZeitAbfragen = False
        SqlCmd = "UPDATE TSESettings SET ZeitVomInternet='0'"
        gdApp.Execute SqlCmd, dbFailOnError
         
    End If
 
Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "chkInterntZeit_Click"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub chkQrCode_Click()
On Error GoTo LOKAL_ERROR

 If chkQrCode.value = vbChecked Then
        MitQrCode = True
        SqlCmd = "UPDATE TSESettings SET mitQR='1'"
        gdApp.Execute SqlCmd, dbFailOnError
        
    Else
    
        MitQrCode = False
        SqlCmd = "UPDATE TSESettings SET mitQR='0'"
        gdApp.Execute SqlCmd, dbFailOnError
         
 End If

Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "chkQrCode_Click"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub chkTSE_Status_Click()
On Error GoTo LOKAL_ERROR

    Dim SqlCmd As String
    
    If chkTSE_Status.value = vbChecked Then
        
        E_TSE_Aktiv = True
        SqlCmd = "UPDATE TSESettings SET TSE_Aktiv='1'"
        gdApp.Execute SqlCmd, dbFailOnError
        HideShowAlleControls True
       
        If Not TSE_OK Then
         TSE_Err = "TSE ist aktiviert aber nicht initialisiert"
        End If
        
    Else
    
        E_TSE_Aktiv = False
        SqlCmd = "UPDATE TSESettings SET TSE_Aktiv='0'"
        gdApp.Execute SqlCmd, dbFailOnError
        HideShowAlleControls False
        
        R_StartTime = ""
        R_FinishTime = ""
        R_TransactionNr = ""
        R_QRCodeAlsText = ""
        R_QRCodeAlsImgPath = ""
        R_FinishSignatur = ""
        R_StartSignatur = ""
        R_FINISH_SIG_Zaehler = 0
        R_START_SIG_Zaehler = 0
    
        TSE_Err = "TSE ist deaktiviert"
       
    End If
     
Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "chkTSE_Status_Click"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command1_Click()

MsgBox ("beim alten Druckmodus wird kein QR-Code gedruckt, " & vbNewLine & "weil manche Bondrucker veraltet sind (unterstützen diese Besonderheit nicht).")

End Sub

Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR

    If Trim(TxtTimeout.Text) <> "" Then
     
          sSQL = "UPDATE TSESettings SET TseTimeOut=" & TxtTimeout.Text
          gdApp.Execute sSQL, dbFailOnError
          
          TSE_TIMEOUT = CInt(TxtTimeout.Text)
          MsgBox ("Timout erfolgreich gespeichert")
          
    Else
          TxtTimeout.Text = TSE_TIMEOUT
          MsgBox ("Tse Timeout ist falsch (muss 1 - 120 Sekunden sein) !!!")
    End If
    

Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
 Check_TSE_Einstellugen
End Sub


Private Sub btnVerbinden_Click()
   
Me.SetFocus
btnVerbinden.Enabled = False

TSE_Initialisieren

btnVerbinden.Enabled = True
btnVerbinden.SetFocus

    
End Sub

Private Sub txtTimeout_Change()
On Error GoTo LOKAL_ERROR

 TxtTimeout.BackColor = vbWhite

 Dim textval As String
  
 textval = Trim(TxtTimeout.Text)
 textval = Replace(textval, ".", "")
 textval = Replace(textval, ",", "")
 
  If IsNumeric(textval) Then
  
       If CInt(textval) >= 1 And CInt(textval) <= 120 Then
            TxtTimeout.Text = CStr(textval)
       Else
            TxtTimeout.Text = ""
       End If
  Else
       TxtTimeout.Text = ""
  End If
  
   
 
Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "TseEinstellungen"
    Fehler.gsFunktion = "txtTimeout_Change"
    Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
  
End Sub
