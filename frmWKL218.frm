VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL218 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TSE-Einstellungen"
   ClientHeight    =   3675
   ClientLeft      =   2055
   ClientTop       =   2865
   ClientWidth     =   7845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3675
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox TextBoxTssClientId 
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
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CheckBox Check4 
      Caption         =   "TSE schreiben"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox TextBoxTssId 
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
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.TextBox TextBoxTssApiSecret 
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
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   6135
   End
   Begin sevCommand3.Command Command2 
      Height          =   615
      Index           =   0
      Left            =   5520
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
   Begin VB.TextBox TextBoxTssApikey 
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
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   0
      Top             =   840
      Width           =   6135
   End
   Begin sevCommand3.Command ButtonTssGenerateIds 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      Caption         =   "TSE Verbindung einrichten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Achtung: Kasse ist nicht mit einer TSE verbunden!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   11
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Client ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   -720
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "TSE ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "API Secret:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "API Key:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmWKL218"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TSE_zeigen(bZeigen As Boolean)
On Error GoTo LOKAL_ERROR

    TextBoxTssApikey.Visible = bZeigen
    TextBoxTssApiSecret.Visible = bZeigen
    TextBoxTssId.Visible = bZeigen
    TextBoxTssClientId.Visible = bZeigen
    
    Label1(0).Visible = bZeigen
    Label1(1).Visible = bZeigen
    Label1(2).Visible = bZeigen
    Label1(3).Visible = bZeigen
    
    ButtonTssGenerateIds.Visible = bZeigen
    Label1(4).Visible = bZeigen
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TSE_zeigen"
    Fehler.gsFehlertext = "Im Programmteil TSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check4_Click()
On Error GoTo LOKAL_ERROR

    If Check4.Value = vbChecked Then
        TSE_zeigen True
    Else
        TSE_zeigen False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check4_Click"
    Fehler.gsFehlertext = "Im Programmteil TSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    Select Case Index
        Case 0      'Speichern
        
            insertTSE TextBoxTssApikey.Text, TextBoxTssApiSecret.Text, TextBoxTssId.Text, TextBoxTssClientId
            Unload frmWKL218
        
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil TSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub insertTSE(sApiKey As String, sApiSecret As String, sTSEID As String, sClientId As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo0 As Integer
   
    If Check4.Value = vbChecked Then
        bo0 = -1
    Else
        bo0 = 0
    End If
    
    sSQL = "Delete from TSE_ONLEINSTELLUNG "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into TSE_ONLEINSTELLUNG ( "
    sSQL = sSQL & " TSE_SCHREIBEN "
    sSQL = sSQL & ", APIKEY "
    sSQL = sSQL & ", APISECRET "
    sSQL = sSQL & ", TSEID "
    sSQL = sSQL & ", CLIENTID "
    sSQL = sSQL & " ) Values "
    sSQL = sSQL & " (" & bo0 & ""
    sSQL = sSQL & " ,'" & sApiKey & "'"
    sSQL = sSQL & " ,'" & sApiSecret & "'"
    sSQL = sSQL & " ,'" & sTSEID & "'"
    sSQL = sSQL & " ,'" & sClientId & "'"
    sSQL = sSQL & " )"
    gdApp.Execute sSQL, dbFailOnError
    
    
    
    
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "insertTSE"
    Fehler.gsFehlertext = "Im Programmteil TSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    LeseTSE_EINSTELLUNG
    
    TextBoxTssApikey.Text = gsTSE_APIKEY
    TextBoxTssApiSecret.Text = gsTSE_APISECRET
    TextBoxTssId.Text = gsTSE_TSEID
    TextBoxTssClientId.Text = gsTSE_CLIENTID
    
    If gbTSE_SCHREIBEN = True Then
        Check4.Value = vbChecked
        TSE_zeigen True
    Else
        Check4.Value = vbUnchecked
        TSE_zeigen False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil TSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub




