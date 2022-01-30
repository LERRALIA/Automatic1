VERSION 5.00
Begin VB.Form frmWKL138 
   Caption         =   "Geschäftsanalyse"
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
   Begin sevCommand3.Command Command5 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Ermitteln"
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
      Index           =   2
      Left            =   9600
      TabIndex        =   7
      Top             =   6600
      Width           =   2055
   End
   Begin VB.ComboBox cboGANATage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin sevCommand3.Command Command4 
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
      TabIndex        =   4
      Top             =   360
      Width           =   345
   End
   Begin sevCommand3.Command Command5 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Drucken"
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
      Index           =   1
      Left            =   9600
      TabIndex        =   3
      Top             =   7200
      Width           =   2055
   End
   Begin sevCommand3.Command Command5 
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
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "vorhandene Auswertungen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3015
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
      TabIndex        =   2
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
      Caption         =   "Geschäftsanalyse"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmWKL138"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
    
    Case 11
        gsHelpstring = "Geschäftsanalyse"
        frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Geschäftsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim lDate As Long
    
Select Case Index
    Case 0
        Unload frmWKL138
    Case 1
        loeschNEW "GANALYSEPRINT", gdBase
        CreateTableT2 "GANALYSEPRINT", gdBase
        
        If Not IsDate(cboGANATage.Text) Then
            Exit Sub
        End If
        
        lDate = CLng(DateValue(cboGANATage.Text))
        
        sSQL = "Insert into GANALYSEPRINT select * from GANALYSEALL where Datum = " & lDate
        gdBase.Execute sSQL, dbFailOnError
        
        reportbildschirm "", "zZEN154"
    Case 2
        
        AnalyseZusammenstellen Label1(4)
        
        fülleGANATage cboGANATage, gcFilNr
        
        Command5_Click 1
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Geschäftsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    fülleGANATage cboGANATage, gcFilNr
    
    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Geschäftsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fülleGANATage(cbox As ComboBox, sFil As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim iFil As Integer
    
    cbox.Clear
    
    cFeld = ""
    
    If sFil = "alle Filialen" Then
        iFil = 0
    Else
        iFil = CInt(Left$(sFil, 3))
    End If

    sSQL = "Select distinct(DATUM) from GANALYSEALL "
    sSQL = sSQL & " where Filiale = " & iFil & " "
    sSQL = sSQL & " order by DATUM desc "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cFeld = ""
            If Not IsNull(rsrs!Datum) Then
                cFeld = rsrs!Datum
                cbox.AddItem cFeld
                
                If cbox.Text = "" Then
                    cbox.Text = cFeld
                End If
            End If
            
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "fülleGANATage"
    Fehler.gsFehlertext = "In der Geschäftsanalyse ist ein Fehler aufgetreten."
    
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

