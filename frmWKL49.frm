VERSION 5.00
Begin VB.Form frmWKL49 
   Caption         =   "Farbmerkmale der Artikel"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   Icon            =   "frmWKL49.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9750
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   9
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   53
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   8
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   51
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   7
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   49
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   6
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   47
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   5
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   45
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   4
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   43
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   41
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   39
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   37
      Top             =   840
      Width           =   3375
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Übernehmen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   10
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   9
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   8
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   6
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   5
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Schließen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   0
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Bevor Sie Datenbankbefehle eingeben sollten Sie zur Sicherheit die Datenbank sichern!!!"
      Height          =   615
      Index           =   30
      Left            =   5040
      TabIndex        =   85
      Top             =   6480
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Datenbankbefehle geben Sie ein unter: SERVICE/DATENBANK/DATENBANKBEFEHL"
      Height          =   615
      Index           =   29
      Left            =   5040
      TabIndex        =   84
      Top             =   5760
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update Artikel set Awm = '2' where  Awm = '3'"
      Height          =   375
      Index           =   28
      Left            =   5040
      TabIndex        =   83
      Top             =   5280
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sie möchten alle Artikel mit dem Farbmerkmal rot auf grün setzen."
      Height          =   615
      Index           =   27
      Left            =   5040
      TabIndex        =   82
      Top             =   4560
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Beispiel eines Datenbankbefehls"
      Height          =   255
      Index           =   26
      Left            =   5040
      TabIndex        =   81
      Top             =   4200
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "92"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   80
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "93"
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   79
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "94"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   78
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "95"
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   77
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "0"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   76
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   75
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   74
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "98"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   73
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "19"
      Height          =   255
      Index           =   17
      Left            =   5040
      TabIndex        =   72
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "18"
      Height          =   255
      Index           =   16
      Left            =   5040
      TabIndex        =   71
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "17"
      Height          =   255
      Index           =   15
      Left            =   5040
      TabIndex        =   70
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "16"
      Height          =   255
      Index           =   14
      Left            =   5040
      TabIndex        =   69
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "15"
      Height          =   255
      Index           =   13
      Left            =   5040
      TabIndex        =   68
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "14"
      Height          =   255
      Index           =   12
      Left            =   5040
      TabIndex        =   67
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "13"
      Height          =   255
      Index           =   11
      Left            =   5040
      TabIndex        =   66
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "12"
      Height          =   255
      Index           =   10
      Left            =   5040
      TabIndex        =   65
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "11"
      Height          =   255
      Index           =   9
      Left            =   5040
      TabIndex        =   64
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   63
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   62
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   61
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   60
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   59
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   58
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   57
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   56
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   55
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   9
      Left            =   5280
      TabIndex        =   54
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   52
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   50
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   48
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   46
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   44
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   42
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   40
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   38
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   17
      Left            =   360
      TabIndex        =   36
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   35
      Top             =   6720
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   34
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   33
      Top             =   6360
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   32
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   31
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   30
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   29
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   27
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   26
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   25
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   24
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   23
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   22
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   21
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   20
      Top             =   4200
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   9600
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   19
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   18
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   16
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Farbmerkmale der Artikel"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmWKL49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            speichern
            Unload frmWKL49
        Case 1
            Unload frmWKL49
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
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    Modul6.alternativFarbform Me, Label1
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    
    For i = 1 To 9
        Label2(i).Caption = " Beispiel"
        Label2(i).BackColor = glfarbe(i)
        Label2(i).ForeColor = vbBlack
    Next i
    
    For i = 1 To 9
        Label5(i).Caption = " Beispiel"
        Label5(i).BackColor = glfarbe2(i)
        Label5(i).ForeColor = vbBlack
    Next i
    
    lesen
    
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
    
    loeschNEW "FARBMERK", gdBase
    CreateTable "FARBMERK", gdBase
    
    
    For i = 1 To 9
        sSQL = "Insert into Farbmerk (FarbText,farbNr)"
        sSQL = sSQL & " values ('" & Text1(i).Text & "', " & i & ")"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Next
    
    For i = 1 To 9
        sSQL = "Insert into Farbmerk (FarbText,farbNr)"
        sSQL = sSQL & " values ('" & Text2(i).Text & "', " & i + 10 & ")"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
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
    
    If Not NewTableSuchenDBKombi("FARBMERK", gdBase) Then
        CreateTable "FARBMERK", gdBase
        For i = 1 To 9
            Text1(i).Text = ""
            Text2(i).Text = ""
        Next
    End If
    
    sSQL = "Select * from farbmerk  where farbnr < 10 order by Farbnr"
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
    
    sSQL = "Select * from farbmerk  where farbnr between 11 and 19 order by Farbnr"
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If Not IsNull(rs!farbtext) Then
                Text2(CInt(rs!FARBNR) - 10).Text = rs!farbtext
            Else
                Text2(CInt(rs!FARBNR) - 10).Text = ""
            End If
        rs.MoveNext
        Loop
    Else
        For i = 1 To 9
            Text2(i).Text = ""
        Next
    End If
    rs.Close: Set rs = Nothing
    
    Label2(10).Caption = " Beispiel"
    Label2(11).Caption = " Beispiel"
    Label2(12).Caption = " Beispiel"
    Label2(13).Caption = " Beispiel"
    Label2(14).Caption = " Beispiel"
    Label2(15).Caption = " Beispiel"
    Label2(16).Caption = " Beispiel"
    Label2(17).Caption = " Beispiel"
    
    Label3(0).Caption = "neuer Artikel"
    Label3(1).Caption = "soeben angefügt"
    Label3(2).Caption = "automatisch kalkulierter"
    Label3(3).Caption = "ohne Farbmerkmal"
    Label3(4).Caption = "nicht geliefert"
    Label3(5).Caption = "für Preisaktion vorgesehen"
    Label3(6).Caption = "befindet sich in Preisaktion"
    Label3(7).Caption = "seit 2 Jahren oder noch nie verkauft"
    
    Label2(10).BackColor = vbWhite
    Label2(11).BackColor = vbWhite
    Label2(12).BackColor = vbYellow
    Label2(13).BackColor = glfarbe(0)
    Label2(14).BackColor = vbBlue
    
    Label2(15).BackColor = glfarbe(0)
    Label2(16).BackColor = vbWhite
    Label2(17).BackColor = vbBlack
    
    
    
    Label2(10).ForeColor = vbRed
    Label2(11).ForeColor = vbBlue
    Label2(12).ForeColor = vbBlue
    Label2(13).ForeColor = vbBlack
    Label2(14).ForeColor = vbBlack
    
    Label2(15).ForeColor = vbBlue
    Label2(16).ForeColor = vbGreen
    Label2(17).ForeColor = vbWhite

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

    If Index < 10 Then
        gsArtikelFarbe = Index
        gsBackcolor = Label2(Index).BackColor
        gsForecolor = Label2(Index).ForeColor
        Command1_Click 1
    End If
    
    If Index = 13 Then
        gsArtikelFarbe = "0"
        gsBackcolor = Label2(Index).BackColor
        gsForecolor = Label2(Index).ForeColor
        Command1_Click 1
    End If
    
    If Index = 10 Then
        gsArtikelFarbe = "98"
        gsBackcolor = Label2(Index).BackColor
        gsForecolor = Label2(Index).ForeColor
        Command1_Click 1
    End If
    
    If Index = 17 Then
        gsArtikelFarbe = "92"
        gsBackcolor = Label2(Index).BackColor
        gsForecolor = Label2(Index).ForeColor
        Command1_Click 1
    End If
    
    If Index = 16 Then
        gsArtikelFarbe = "93"
        gsBackcolor = Label2(Index).BackColor
        gsForecolor = Label2(Index).ForeColor
        Command1_Click 1
    End If
    
    If Index = 15 Then
        gsArtikelFarbe = "94"
        gsBackcolor = Label2(Index).BackColor
        gsForecolor = Label2(Index).ForeColor
        Command1_Click 1
    End If
    
    If Index = 14 Then
        gsArtikelFarbe = "95"
        gsBackcolor = Label2(Index).BackColor
        gsForecolor = Label2(Index).ForeColor
        Command1_Click 1
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label2_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
    Fehlermeldung1
End Sub

Private Sub Label5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index < 10 Then
        gsArtikelFarbe = Index + 10
        gsBackcolor = Label5(Index).BackColor
        gsForecolor = Label5(Index).ForeColor
        Command1_Click 1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label5_Click"
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
Private Sub Text2_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text2(Index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf. "
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text2(Index).BackColor = glSelBack1
    Text2(Index).SelStart = 0
    Text2(Index).SelLength = Len(Text2(Index).Text)
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf."
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case 1 To 9
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%/!"
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case 1 To 9
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%/!"
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmal auf."
    
    Fehlermeldung1
End Sub
