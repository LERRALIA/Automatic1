VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL100 
   Caption         =   "Gutscheinsuche"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.OptionButton Option1 
      Caption         =   "Gutschein Nr"
      CausesValidation=   0   'False
      Height          =   255
      Index           =   5
      Left            =   9480
      TabIndex        =   24
      Tag             =   "gutschnr"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Bediener"
      CausesValidation=   0   'False
      Height          =   255
      Index           =   4
      Left            =   9480
      TabIndex        =   22
      Tag             =   "bednu"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Filiale"
      Height          =   255
      Index           =   3
      Left            =   9480
      TabIndex        =   21
      Tag             =   "filiale"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Kundennummer"
      Height          =   255
      Index           =   2
      Left            =   9480
      TabIndex        =   20
      Tag             =   "kundnr"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Wert"
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   19
      Tag             =   "wert"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ausgabedatum"
      Height          =   255
      Index           =   0
      Left            =   9480
      TabIndex        =   18
      Tag             =   "dat_ausg desc"
      Top             =   1680
      Width           =   1815
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   16
      Top             =   6600
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
      Caption         =   "Suche"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   360
      TabIndex        =   15
      Top             =   1560
      Width           =   9135
   End
   Begin VB.PictureBox picprogress 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   9075
      TabIndex        =   14
      Top             =   7440
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Combo5 
      Height          =   330
      Left            =   8640
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox Combo4 
      Height          =   330
      Left            =   6600
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   330
      Left            =   4560
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   330
      Left            =   2280
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   480
      Width           =   1935
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   1
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
      Caption         =   "Zurück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
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
      Caption         =   "Übernehmen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   9135
   End
   Begin VB.Label Label7 
      Caption         =   "Sortierung"
      Height          =   255
      Left            =   9480
      TabIndex        =   23
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000C&
      Caption         =   "Ausgabedatum"
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
      Left            =   120
      TabIndex        =   17
      Top             =   7920
      Width           =   9135
   End
   Begin VB.Label Label5 
      Caption         =   "Filiale"
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Bediener"
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Ausgabedatum"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Wert"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Kunde"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmWKL100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Combo1_Click()
On Error GoTo LOKAL_ERROR
    
   SucheDaten
      
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Click"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo2_Click()
On Error GoTo LOKAL_ERROR
    
   SucheDaten
      
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_Click"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo3_Click()
On Error GoTo LOKAL_ERROR
    
   SucheDaten
      
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_Click"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo4_Click()
On Error GoTo LOKAL_ERROR
    
   SucheDaten
      
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo4_Click"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo5_Click()
On Error GoTo LOKAL_ERROR
    
   SucheDaten
      
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo5_Click"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case index
        Case Is = 0
            prot_GZ ("SUCHE")
            SucheDaten
        Case Is = 1     'zurück
        prot_GZ ("ZURÜCK")
            gLGutschnum = -1
            Unload frmWKL100
        Case Is = 2 'Übernehmen
            prot_GZ ("ÜBERNEHMEN")
            Dim bFound As Boolean
            Dim lcount As Long
            
            bFound = False
            
            If List2.ListCount = 0 Then
                Exit Sub
            End If
            
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    bFound = True
                    Exit For
                End If
            Next lcount
            
            If bFound Then
              '  gLGutschnum = CLng(Left(List2.list(lcount), 8))  ' <<<< 25.02.2022 VL
                gLGutschnum = CLng(Left(List2.list(lcount), 10))
            Else
                
                Exit Sub
            End If
            Unload frmWKL100
        
    End Select
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheDaten()
    On Error GoTo LOKAL_ERROR
    Dim cWhe As String
    
    
    
    If gbKL_LIVEGUTSCHEIN Then
    
        cWhe = ""
    
        If Combo1.Text <> "alle Werte" Then
            cWhe = cWhe & " and Wert = " & SwapStr(Combo1.Text, ",", ".")
        End If
        
        If Combo2.Text <> "alle Kunden" Then
            cWhe = cWhe & " and AUSG_KUNDNR = " & Combo2.Text
        End If
        
        If Combo3.Text <> "alle" Then
        
            Dim lHeute As Long

            lHeute = Fix(DateValue(Combo3.Text))
            lHeute = lHeute - 2
            cWhe = cWhe & " and AUSG_DATUM = " & Trim$(Str$(lHeute)) & " "

'            cWhe = cWhe & " and AUSG_DATUM = #29.05.2017# "
            
'            cWhe = cWhe & " and AUSG_DATUM = '" & DateValue(Combo3.Text) & "' "
        End If
        
        If Combo4.Text <> "alle Bediener" Then
            cWhe = cWhe & " and AUSG_Bediener = " & Left(Combo4.Text, 4)
        End If
        
        If Combo5.Text <> "alle Filialen" Then
            cWhe = cWhe & " and AUSG_FILIALE = " & Left(Combo5.Text, 3)
        End If
    
    
        LeseOffeneGutscheineWK100_KL_SQL cWhe
    Else
    
    
        cWhe = ""
    
        If Combo1.Text <> "alle Werte" Then
            cWhe = cWhe & " and Wert = " & SwapStr(Combo1.Text, ",", ".")
        End If
        
        If Combo2.Text <> "alle Kunden" Then
            cWhe = cWhe & " and KUNDNR = " & Combo2.Text
        End If
        
        If Combo3.Text <> "alle" Then
            cWhe = cWhe & " and dat_ausg = " & CLng(Combo3.Text)
        End If
        
        If Combo4.Text <> "alle Bediener" Then
            cWhe = cWhe & " and bednu = " & Left(Combo4.Text, 4)
        End If
        
        If Combo5.Text <> "alle Filialen" Then
            cWhe = cWhe & " and FILIALE = " & Left(Combo5.Text, 3)
        End If
    
    
        LeseOffeneGutscheineWK100 cWhe
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDaten"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
'    WKL100Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Farbform Me, Nothing
    
    Screen.MousePointer = 11
    
    If gbKL_LIVEGUTSCHEIN Then
    
        If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
            'alles okay
            
        Else
        
            schreibeProtokollVPNTXT "Unterbrechung"
        
            Dim sTemp As String
            sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
            sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
        
            MsgBox sTemp, vbCritical + vbOKOnly, "Gutschein-Datenbank nicht erreichbar"
            Screen.MousePointer = 0
            
            Command1(0).Enabled = False
            Exit Sub
        End If
    
    
        fülleSpalte_KL Combo2, "AUSG_KUNDNR", "GUTSCHEINE", "AUSG_KUNDNR", "alle Kunden", "", " where (EINL_DATUM is null or EINL_DATUM = 0 )"
        fülleSpalte_KL Combo1, "WERT", "GUTSCHEINE", "WERT", "alle Werte", "d", " where (EINL_DATUM is null or EINL_DATUM = 0 )"
        fülleSpalte_KL Combo3, "AUSG_DATUM", "GUTSCHEINE", "AUSG_DATUM", "alle", "D", " where (EINL_DATUM is null or EINL_DATUM = 0 )"
'        Combo3.Visible = False
    Else
        fülleSpalte Combo2, "KUNDNR", "GUTSCH", "KUNDNR", "alle Kunden", ""
        fülleSpalte Combo1, "WERT", "GUTSCH", "WERT", "alle Werte", "d"
        fülleSpalte Combo3, "dat_ausg", "GUTSCH", "dat_ausg", "alle", "D"
    End If

    füllefillo Combo5
    filcboBediener Combo4
    
    Screen.MousePointer = 0
    
    Option1(LeselastOptionEinstellung("E100X")).value = True
    
    SucheDaten
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub füllefillo(cbox As ComboBox)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cSatz As String
    Dim cFeld As String
    
    cbox.Clear
    cbox.AddItem "alle Filialen"
    cbox.Text = "alle Filialen"
    
    sSQL = "Select * from filialen order by filialnr"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALNR) Then
            
                cFeld = rsrs!FILIALNR
                cSatz = cSatz & Space(3 - Len(cFeld)) & cFeld
                
                If Not IsNull(rsrs!FILIALNAME) Then
                    cFeld = rsrs!FILIALNAME
                    cSatz = cSatz & Space(2) & cFeld
                    cbox.AddItem cSatz
                    
                    If gcFilNr = Trim(Left(cSatz, 3)) Then
                        cbox.Text = cSatz
                    End If
                End If
            End If
            cSatz = ""
            cFeld = ""
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "füllefillo"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseOffeneGutscheineWK100(cwhere As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim dWert       As Double
    Dim rsrs        As Recordset
    Dim lcount      As Long
    Dim lcountall   As Long
    Dim j           As Integer
    Dim corder      As String
    
    Screen.MousePointer = 11
    
    List1.Clear
    List2.Clear
    
    For j = 0 To 5
        If Option1(j).value = True Then
            corder = Option1(j).Tag
            Exit For
        End If
    Next j
    
    List1.AddItem " Gutsch. Ausgabe am       Wert      Kunde   Bediener    Filiale"
    
    sSQL = "Select * from gutsch where Status <> 'L' and not Wert  is null " & cwhere & " order by  " & corder


    picprogress.Visible = True

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        lcountall = lcount
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            lcount = lcount - 1
            
            j = lcount Mod 3000
            If j = 0 Then
                txtStatus.Text = CStr(lcount * 100 / lcountall)
            Else
                
            End If
            
            If IsNull(rsrs!DAT_EINL) Or rsrs!DAT_EINL = 0 Then
                If Not IsNull(rsrs!gutschnr) Then
                    cFeld = rsrs!gutschnr
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                'cFeld = Space$(8 - Len(cFeld)) & cFeld    '<<<<< 25.02.2022 VL
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cFeld & " "
                
                If Not IsNull(rsrs!DAT_AUSG) Then
                    dWert = rsrs!DAT_AUSG
                Else
                    dWert = 0
                End If
                If dWert > 0 Then
                    cFeld = Format$(dWert, "DD.MM.YYYY")
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cFeld = cFeld & Space$(10 - Len(cFeld))
                cLBSatz = cLBSatz & cFeld & " "
                
                If Not IsNull(rsrs!Wert) Then
                    dWert = rsrs!Wert
                Else
                    dWert = 0
                End If
                cFeld = Format$(dWert, "######0.00")
                cFeld = Trim$(cFeld)
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                
                If Not IsNull(rsrs!Kundnr) Then
                    cFeld = rsrs!Kundnr
                Else
                    cFeld = "0"
                End If
                
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                
                If Not IsNull(rsrs!BEDNU) Then
                    cFeld = rsrs!BEDNU
                Else
                    cFeld = "0"
                End If
                
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                
                If Not IsNull(rsrs!FILIALE) Then
                    cFeld = rsrs!FILIALE
                Else
                    cFeld = "0"
                End If
                
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                
                List2.AddItem cLBSatz
                
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    picprogress.Visible = False
    
    Label6.Caption = "vorhandene Gutscheine: " & lcountall
    Label6.Refresh
    
    Screen.MousePointer = 0
    
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseOffeneGutscheineWK100"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1

    
End Sub
Private Sub LeseOffeneGutscheineWK100_KL_SQL(cwhere As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim dWert       As Double
    Dim rsrs        As Recordset
    Dim lcount      As Long
    Dim lcountall   As Long
    Dim j           As Integer
    Dim corder      As String
    
    Screen.MousePointer = 11
    
    List1.Clear
    List2.Clear
    
    For j = 0 To 5
        If Option1(j).value = True Then
            corder = Option1(j).Tag
            Exit For
        End If
    Next j
    
    Select Case corder
        Case "dat_ausg desc"
            corder = "AUSG_DATUM"
        Case "filiale"
            corder = "AUSG_FILIALE"
        Case "bednu"
            corder = "AUSG_BEDIENER"
        Case "kundnr"
            corder = "AUSG_KUNDNR"
    End Select
    
    List1.AddItem " Gutsch. Ausgabe am       Wert      Kunde   Bediener    Filiale"
    
    Dim stConnect As String
    
    If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
        'alles okay
        
    Else
    
        schreibeProtokollVPNTXT "Unterbrechung"
        
        Dim sTemp As String
        sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
        sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
    
        MsgBox sTemp, vbCritical + vbOKOnly, "Gutschein-Datenbank nicht erreichbar"
        Exit Sub
    End If
    
    
    If gsKL_DSN <> "" Then
        stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
    Else
        stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
    End If
    
    
    
    
    
    Dim dbEAN As DAO.Database
    Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
    
    sSQL = "Select * from GUTSCHEINE "
    sSQL = sSQL & " where (EINL_DATUM is null or EINL_DATUM = 0 ) "
    sSQL = sSQL & " " & cwhere & " order by  " & corder
    Set rsrs = dbEAN.OpenRecordset(sSQL)
    
    picprogress.Visible = True

    If Not rsrs.EOF Then
    
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        lcountall = lcount
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            lcount = lcount - 1
            
            j = lcount Mod 3000
            If j = 0 Then
                txtStatus.Text = CStr(lcount * 100 / lcountall)
            Else
                
            End If
            
            
            If Not IsNull(rsrs!gutschnr) Then
                cFeld = rsrs!gutschnr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            'cFeld = Space$(8 - Len(cFeld)) & cFeld   '<<<< 25.02.2022  VL
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!AUSG_DATUM) Then
                dWert = rsrs!AUSG_DATUM
            Else
                dWert = 0
            End If
            If dWert > 0 Then
                cFeld = Format$(dWert, "DD.MM.YYYY")
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(10 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!Wert) Then
                dWert = rsrs!Wert
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "######0.00")
            cFeld = Trim$(cFeld)
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!AUSG_Kundnr) Then
                cFeld = rsrs!AUSG_Kundnr
            Else
                cFeld = "0"
            End If
            
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!AUSG_BEDIENER) Then
                cFeld = rsrs!AUSG_BEDIENER
            Else
                cFeld = "0"
            End If
            
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!AUSG_FILIALE) Then
                cFeld = rsrs!AUSG_FILIALE
            Else
                cFeld = "0"
            End If
            
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            List2.AddItem cLBSatz
                
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dbEAN.Close
    
    picprogress.Visible = False
    
    Label6.Caption = "vorhandene Gutscheine: " & lcountall
    Label6.Refresh
    
    Screen.MousePointer = 0
    
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseOffeneGutscheineWK100_KL_SQL"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1

    
End Sub

Private Sub Option1_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    speicherlastOptionEinstellung index, "E100X"
'    SucheDaten
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtStatus_Change()
On Error GoTo LOKAL_ERROR

    Dim nProz As Long
  
    nProz = Val(txtStatus.Text)
    ShowProgress picprogress, nProz, 0, 100, True
    picprogress.Refresh
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtstatus_Change"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
