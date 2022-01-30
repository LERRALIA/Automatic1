VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL129 
   Caption         =   "Email"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL129.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chk_Kopie 
      Caption         =   "Kopie an die eigene Adresse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9600
      TabIndex        =   10
      Top             =   6360
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   3240
      Width           =   9255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   5295
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   4
      Top             =   7200
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
      Caption         =   "Senden"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   1
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   11
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      Caption         =   "löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Emailadresse:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Emailtext:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Betreff:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   7920
      Width           =   9255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmWKL129"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Emailadresse_del()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
        
    sSQL = "Delete * from email where adresse = '" & Combo1.Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    Combo1.Text = ""
    fuelleEmail
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Emailadresse_del"
    Fehler.gsFehlertext = "Im Programmteil Email ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL129
        Case 1
            If InStr(1, Combo1.Text, "@") = 0 Or InStr(1, Combo1.Text, ".") = 0 Then
                anzeige "rot2", "Bitte geben Sie die Emailadresse richtig ein!", Label1(4)
            Else
            
                Dim sThema As String
                sThema = ""
                
                If gcBestellEmail.Subject = "Frage/Stammdatenpflege" Then
                    sThema = "Stammdatenpflege"
                End If
                InsertEMAIL Combo1.Text, sThema
                
                sendeMail
                fuelleEmail
            End If
        Case 2 'Email löschen
            Emailadresse_del
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Email ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub

Private Sub Form_Activate()
On Error GoTo LOKAL_ERROR


Text1(1).SetFocus

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Activate"
    Fehler.gsFehlertext = "Im Programmteil Email ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    fuelleEmail
    
    If gcBestellEmail.Subject <> "" Then
        Text1(0).Text = gcBestellEmail.Subject
    End If
    
    If gcBestellEmail.Message <> "" Then
        Text1(1).Text = gcBestellEmail.Message
    End If
    
    If gcBestellEmail.Recipient <> "" Then
        Combo1.Text = gcBestellEmail.Recipient
    End If
    
    
    

    
    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Email ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub fuelleEmail()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim sSQL As String
    
    Combo1.Clear
    
    sSQL = "select * from email"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!adresse) Then
                Combo1.AddItem rsrs!adresse
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelleEmail"
    Fehler.gsFehlertext = "Im Programmteil Email ist ein Fehler aufgetreten."
    
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
Private Sub sendeMail()
    On Error GoTo LOKAL_ERROR
    
    gcBestellEmail.SenderName = ermFirmenBez
    gcBestellEmail.ReplyTo = ermFirmenMail
    gcBestellEmail.SenderEMail = ermFirmenMail '"bestsend@kisswws.de"
    
'    gbCCfromBestlief = True
'    If gbCCfromBestlief = True Then
'        gcBestellEmail.CC = ermFirmenMail
'    End If
    
    
    If chk_Kopie.Value = vbChecked Then
        gcBestellEmail.CC = ermFirmenMail
    End If
    
    gcBestellEmail.Recipient = Combo1.Text
    gcBestellEmail.SMTPAUTH = True
    
    'die 5 Werte sind jetzt in der Tabelle Kassein
    
    gcBestellEmail.ServerName = gcSMTP_SERVER
    gcBestellEmail.ServerPort = gcSMTP_PORT
    gcBestellEmail.Username = gcSMTP_USER
    gcBestellEmail.Password = gcSMTP_PW
    gcBestellEmail.SSL = gbSMTP_SSL
    

    
    'Betreff
    gcBestellEmail.Subject = Text1(0).Text
    gcBestellEmail.Message = Text1(1).Text

    gcBestellEmail.AutoZIP = False
'    gcBestellEmail.Attachment1 = ""
'    gcBestellEmail.Attachment2 = ""
    gcBestellEmail.Attachment3 = ""
    gcBestellEmail.Attachment4 = ""
    gcBestellEmail.Attachment5 = ""

    'Ab hier an frmwkl38 übergeben

'    giKissFtpMode = 23 ' Email verschicken
    giKissFtpMode = 46 ' Email verschicken neu SSL
    frmWKL38.Show 1

    Pause (1)
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "sendeMail"
    Fehler.gsFehlertext = "Im Programmteil Email ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Text1(Index).SelStart = Len(Text1(Index).Text)
    Text1(Index).SelLength = 0
    Text1(Index).BackColor = glSelBack1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Email ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command5_Click 1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kunde suchen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Email ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
