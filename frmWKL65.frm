VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL65 
   Caption         =   "Farbmerkmale der Kunden"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   Icon            =   "frmWKL65.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   4800
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   0
      Top             =   5880
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
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   10
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   9
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   8
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   7
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   6
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   5
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   4
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   3
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   2
      Top             =   3960
      Width           =   3495
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
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
      Caption         =   "‹bernehmen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "Klicken Sie auf die Farbe!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Farbmerkmale der Kunden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   11
      Top             =   5280
      Width           =   3495
   End
End
Attribute VB_Name = "frmWKL65"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    
    Screen.MousePointer = 11
    
    Modul6.alternativFarbform Me, Label1
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    
    For i = 1 To 9
        Label2(i).Caption = "Beispiel"
        Label2(i).BackColor = glfarbe(i)
        Label2(i).ForeColor = vbBlack
    Next i
    
    lesen
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
    Fehlermeldung1
End Sub
Private Sub speichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim i As Integer
    
    loeschNEW "FARBKU", gdBase
    CreateTable "FARBKU", gdBase
    
    For i = 1 To 9
        sSQL = "Insert into FARBKU (FarbText,farbNr)"
        sSQL = sSQL & " values ('" & Text1(i).Text & "', " & i & ")"
        gdBase.Execute sSQL, dbFailOnError
    Next
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichern"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
    Fehlermeldung1
End Sub
Private Sub lesen()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim sSQL As String
    Dim i As Integer
    
    If Not NewTableSuchenDBKombi("FARBKU", gdBase) Then
        CreateTable "FARBKU", gdBase
        For i = 1 To 9
            Text1(i).Text = ""
        Next
    End If
    
    sSQL = "Select * from FARBKU  where farbnr < 10 order by Farbnr"
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If Not IsNull(rs!farbtext) Then
                Text1(rs!FARBNR).Text = rs!farbtext
            Else
                Text1(rs!FARBNR).Text = ""
            End If
        rs.MoveNext
        Loop
    Else
        For i = 1 To 9
            Text1(i).Text = ""
        Next
    End If
    rs.Close: Set rs = Nothing
    
'    Label2(10).Caption = "Beispiel"
'    Label2(11).Caption = "Beispiel"
'    Label2(12).Caption = "Beispiel"
    Label2(13).Caption = "Beispiel"
    
'    Label3(0).Caption = "neuer Artikel"
'    Label3(1).Caption = "soeben angef¸gt"
'    Label3(2).Caption = "automatisch kalkulierter"
    Label3(3).Caption = "ohne Farbmerkmal"
    
'    Label2(10).BackColor = vbWhite
'    Label2(11).BackColor = vbWhite
'    Label2(12).BackColor = vbYellow
    Label2(13).BackColor = glfarbe(0)
    
'    Label2(10).ForeColor = vbRed
'    Label2(11).ForeColor = vbBlue
'    Label2(12).ForeColor = vbBlue
    Label2(13).ForeColor = vbBlack
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lesen"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
    Fehlermeldung1
End Sub

Private Sub Label2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    gsBackcolor = Label2(Index).BackColor
    gsForecolor = Label2(Index).ForeColor
    
    If Index = 13 Then
        gsKundenFarbe = "0"
    Else
        gsKundenFarbe = Index
    End If
    Command1_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "label2_click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
    Fehlermeldung1
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
'        Case 0
'
'            Unload frmWKL65
        Case 1
            speichern
            Unload frmWKL65
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
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

