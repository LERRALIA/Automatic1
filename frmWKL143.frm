VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL143 
   Caption         =   "Artikel Strichcodes"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL143.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.OptionButton Option1 
      Caption         =   "6 Stellen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   9600
      TabIndex        =   10
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "8 Stellen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9600
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "12 Stellen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "13 Stellen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9600
      TabIndex        =   7
      Top             =   1920
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox Text1 
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
      Index           =   0
      Left            =   120
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   9255
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   6
      Left            =   9600
      TabIndex        =   3
      Top             =   7080
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
      Caption         =   "neu erstellen"
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
      Caption         =   "Schlieﬂen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   9255
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
      Caption         =   "Artikel Strichcodes 13 Stellen"
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
      Width           =   11535
   End
End
Attribute VB_Name = "frmWKL143"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR


    Select Case Index
        Case 0
            Unload frmWKL143
        Case 6
        
            Text1(0).Text = ""
            If Option1(0).Value = True Then
                Strichcodeliste 13, 7
                fuellelisteStrichcodes "", 13, 7
            ElseIf Option1(1).Value = True Then
                Strichcodeliste 12, 6
                fuellelisteStrichcodes "", 12, 6
            ElseIf Option1(2).Value = True Then
                Strichcodeliste 8, 3
                fuellelisteStrichcodes "", 8, 3
            ElseIf Option1(3).Value = True Then
                Strichcodeliste 6, 2
                fuellelisteStrichcodes "", 6, 2
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Strichcodes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Strichcodeliste(ilen As Integer, iLiefteil As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "bitte warten, Daten werden ermittelt...", Label1(4)
    
    loeschNEW "DEAN" & ilen, gdBase
    
    sSQL = "Select Left(ean," & iLiefteil & ") as kean,linr into DEAN" & ilen & " from artikel "
    sSQL = sSQL & " where len(ean) = " & ilen
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into DEAN" & ilen & " Select Left(ean2," & iLiefteil & ") as kean,linr from artikel "
    sSQL = sSQL & " where len(ean2) = " & ilen
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into DEAN" & ilen & " Select Left(ean3," & iLiefteil & ") as kean,linr from artikel "
    sSQL = sSQL & " where len(ean3) = " & ilen
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "KEAN" & ilen, gdBase
    
    sSQL = "Select distinct(kean) as kurzean,linr,'' as LIEFBEZ into KEAN" & ilen & " from DEAN" & ilen
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "update KEAN" & ilen & " inner join Lisrt on kean" & ilen & ".linr = lisrt.linr set kean" & ilen & ".Liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    If ilen = 13 Then
        sSQL = "Delete from  KEAN" & ilen & " where val(kurzean) < 100000 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    loeschNEW "DEAN" & ilen, gdBase
    
    Screen.MousePointer = 0
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Strichcodeliste"
    Fehler.gsFehlertext = "Im Programmteil Artikel Strichcodes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Sub fuellelisteStrichcodes(sStrich As String, ilen As Long, ilenLiefbez As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lAnz As Long
    Dim cFeld As String
    
    Screen.MousePointer = 11
    
    
    
    List1.Clear
    List2.Clear
    List2.AddItem ilenLiefbez & " Stellen Lieferantenbezeichnung              Linr"
    
    sSQL = "Select * from KEAN" & ilen
    If sStrich <> "" Then
        sSQL = sSQL & " where kurzean like  '" & sStrich & "*'"
    End If
    sSQL = sSQL & " order by KURZEAN "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Kurzean) Then
                cFeld = rsrs!Kurzean
            Else
                cFeld = ""
            End If
            
            cLBSatz = Space(9 - Len(cFeld)) & cFeld & " "
            
            If Not IsNull(rsrs!LIEFBEZ) Then
                cFeld = rsrs!LIEFBEZ
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))

            If Not IsNull(rsrs!linr) Then
                cFeld = rsrs!linr
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & Space(6 - Len(cFeld)) & cFeld & " "
            
            List1.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

    If lAnz = 0 Then
        anzeige "rot2", "Keine Lieferanten gefunden.", Label1(4)
    ElseIf lAnz = 1 Then
        anzeige "normal", lAnz & " Lieferantenkennung aus allen Strichcodes(L‰nge " & ilen & ") wurde ermittelt.", Label1(4)
    Else
        anzeige "normal", lAnz & " verschiedene Lieferantenkennungen aus allen Strichcodes(L‰nge " & ilen & ") wurde ermittelt.", Label1(4)
    End If
    

    Screen.MousePointer = 0
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellelisteStrichcodes"
    Fehler.gsFehlertext = "Im Programmteil Artikel Strichcodes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    If NewTableSuchenDBKombi("KEAN13", gdBase) = False Then
        Strichcodeliste 13, 7
        
    End If
    fuellelisteStrichcodes "", 13, 7
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Strichcodes ist ein Fehler aufgetreten."
    
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



Private Sub Option1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Option1(0).Value = True Then
        lblUeberschrift.Caption = "Artikel Strichcodes 13 Stellen"
        lblUeberschrift.Refresh
        
        If NewTableSuchenDBKombi("KEAN13", gdBase) = False Then
            Strichcodeliste 13, 7
        End If
        fuellelisteStrichcodes Text1(0).Text, 13, 7
    ElseIf Option1(1).Value = True Then
        lblUeberschrift.Caption = "Artikel Strichcodes 12 Stellen"
        lblUeberschrift.Refresh
        If NewTableSuchenDBKombi("KEAN12", gdBase) = False Then
            Strichcodeliste 12, 6
        End If
        fuellelisteStrichcodes Text1(0).Text, 12, 6
    ElseIf Option1(2).Value = True Then
        lblUeberschrift.Caption = "Artikel Strichcodes 8 Stellen"
        lblUeberschrift.Refresh
        If NewTableSuchenDBKombi("KEAN8", gdBase) = False Then
            Strichcodeliste 8, 3
        End If
        fuellelisteStrichcodes Text1(0).Text, 8, 3
    ElseIf Option1(3).Value = True Then
        lblUeberschrift.Caption = "Artikel Strichcodes 6 Stellen"
        lblUeberschrift.Refresh
        If NewTableSuchenDBKombi("KEAN6", gdBase) = False Then
            Strichcodeliste 6, 2
        End If
        fuellelisteStrichcodes Text1(0).Text, 6, 2
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Option1(0).Value = True Then
        If Len(Text1(0).Text) > 0 Then
            fuellelisteStrichcodes Text1(0).Text, 13, 7
        End If
    ElseIf Option1(1).Value = True Then
        If Len(Text1(0).Text) > 0 Then
            fuellelisteStrichcodes Text1(0).Text, 12, 6
        End If
    ElseIf Option1(2).Value = True Then
        If Len(Text1(0).Text) > 0 Then
            fuellelisteStrichcodes Text1(0).Text, 8, 3
        End If
    ElseIf Option1(3).Value = True Then
        If Len(Text1(0).Text) > 0 Then
            fuellelisteStrichcodes Text1(0).Text, 6, 2
        End If
    End If
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Artikel Strichcodes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
