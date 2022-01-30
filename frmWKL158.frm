VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL158 
   Caption         =   "Schwerpunkt"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL158.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   7680
      TabIndex        =   9
      Top             =   1320
      Width           =   3975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   7680
      TabIndex        =   8
      Top             =   4320
      Width           =   3975
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3975
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   3975
   End
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   3
      Top             =   360
      Width           =   345
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
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
      Caption         =   "?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
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
   Begin VB.Label Label1 
      Caption         =   "Möchten Sie alle Kunden sehen?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   120
      MouseIcon       =   "frmWKL158.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   5280
      TabIndex        =   12
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Schwerpunkt/Marke"
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
      Left            =   7680
      TabIndex        =   11
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Schwerpunkt/Produktgruppe"
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
      Left            =   7680
      TabIndex        =   10
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Schwerpunkt/Lieferant"
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
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Schwerpunkt/Artikelgruppe"
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
      Index           =   12
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Schwerpunkt"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmWKL158"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lLinr As Long
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL158
        
        Case 11
            gsHelpstring = "Schwerpunkt"
            frmWKL110.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Schwerpunkt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    
    
    lLinr = ZeigeSchwerpunktLinr(gckundnr, List4)
    If lLinr > 0 Then
        Label1(3).Caption = "Möchten Sie alle Kunden sehen, die Ihren Schwerpunkt von über 80% in " & ermLiefBez(lLinr) & " sehen?"
        Label1(3).Visible = True
    End If
    ZeigeSchwerpunktAGN gckundnr, List5
    ZeigeSchwerpunktPGN gckundnr, List1
    ZeigeSchwerpunktMarke gckundnr, List2
    
    Label1(19).Caption = WhatIsXfromKu(gckundnr, "vorName") & " " & WhatIsXfromKu(gckundnr, "Name")
    
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Schwerpunkt ist ein Fehler aufgetreten."
    
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
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 3
            alleKundenLinrSchwerpunkt lLinr, 80
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Schwerpunkt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub alleKundenLinrSchwerpunkt(lLinr As Long, dProzente As Double)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cKundnr As String
    Dim dSumme As Double
    Dim dAnteil As Double
    Dim dProz As Double
    
    
    loeschNEW "SchwerKundnr", gdBase
    CreateTableT2 "SCHWERKUNDNR", gdBase
    
    Screen.MousePointer = 11
    
    sSQL = "Insert into SchwerKundnr Select Kundnr, 0.0 as summe, sum(preis) as Anteil ,0.0 as Proz  from kassjour where linr = " & lLinr
    sSQL = sSQL & " and preis > 0 "
    sSQL = sSQL & " and Kundnr > 0 "
    sSQL = sSQL & " group by kundnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into SchwerKundnr Select Kundnr, 0.0 as summe, sum(preis) as Anteil ,0.0 as Proz  from kundkass "
    sSQL = sSQL & " inner join artlief on kundkass.artnr = artlief.artnr "
    sSQL = sSQL & " Where artlief.linr = " & lLinr
    sSQL = sSQL & " and preis > 0 "
    sSQL = sSQL & " and Kundnr > 0 "
    sSQL = sSQL & " group by kundnr "
    gdBase.Execute sSQL, dbFailOnError
    
'    Kundkass dazu
    
    Set rsrs = gdBase.OpenRecordset("SchwerKundnr")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!Kundnr) Then
                cKundnr = rsrs!Kundnr
            Else
                cKundnr = "0"
            End If
            
            If Not IsNull(rsrs!Anteil) Then
                dAnteil = rsrs!Anteil
            Else
                dAnteil = "0"
            End If
            
            If Val(cKundnr) > 0 Then
                dSumme = ermumsatzTotal(cKundnr, False)
            End If
            
            If dSumme <> 0 Then
                dProz = dAnteil * 100 / dSumme
            End If
            
            rsrs.Edit
            rsrs!Summe = dSumme
            rsrs!Proz = dProz
            rsrs.Update
    
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    sSQL = "Delete from  SchwerKundnr where PROZ  < " & dProzente
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "alleKundenLinrSchwerpunkt"
    Fehler.gsFehlertext = "Im Programmteil Schwerpunkt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
