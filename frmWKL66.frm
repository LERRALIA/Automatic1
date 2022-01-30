VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL66 
   Caption         =   "Bankverbindung"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmWKL66.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   4680
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   120
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2160
      Width           =   3495
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   5040
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
      Caption         =   "Übernehmen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   0
      Top             =   5040
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "ORT"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   12
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "PLZ"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Bank"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Vorname"
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
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
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
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Bankleitzahl"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Kontonummer"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Bankverbindung"
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
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Kundnr"
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "frmWKL66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            speichern
            Unload frmWKL66
        Case 1
            
            Unload frmWKL66
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    Fehlermeldung1
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    
    Screen.MousePointer = 11
    
    Modul6.alternativFarbform Me, Label1
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    
    
    
    lesen
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    Fehlermeldung1
End Sub
Private Sub speichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Screen.MousePointer = 11

    
    sSQL = "Delete from  BANKKU where kundnr = " & gckundnr
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BANKKU (KUNDNR,BLZ,KTNR)"
    sSQL = sSQL & " values ('" & gckundnr & "','" & Text1(2).Text & "','" & Text1(1).Text & "')"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichern"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    Fehlermeldung1
End Sub
Private Sub lesen()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim sSQL As String
    Dim i As Integer
    
    If gckundnr = "" Then
        Exit Sub
    End If
    
    Label2(3).Caption = ""
    Label2(4).Caption = ""
    Label2(0).Caption = ""
               
    
    Label4(0).Caption = gckundnr
    Label4(2).Caption = gckuVorname
    Label4(1).Caption = gckuname
    
    
    
    sSQL = "Select * from BANKKU  where KUNDNR = " & gckundnr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!KTNR) Then
            Text1(1).Text = rs!KTNR
        Else
            Text1(1).Text = ""
        End If

        If Not IsNull(rs!BLZ) Then
            Text1(2).Text = rs!BLZ
        Else
            Text1(2).Text = ""
        End If

    Else
        For i = 1 To 2
            Text1(i).Text = ""
        Next
    End If
    rs.Close: Set rs = Nothing
    

    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lesen"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    Fehlermeldung1
End Sub

Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim sSQL As String
    Dim i As Integer
    
    If Index = 2 Then
        If Len(Text1(2).Text) = 8 Then
            sSQL = "Select * from BANKen  where BLZ = '" & Trim(Text1(2).Text) & "'"
            Set rs = gdBase.OpenRecordset(sSQL)
            If Not rs.EOF Then
            
                If Not IsNull(rs!BankName) Then
                    anzeige "normal", rs!BankName, Label2(0)
                    
                    If Not IsNull(rs!Plz) Then
                        Label2(3).Caption = rs!Plz
                    End If
                    
                    If Not IsNull(rs!Ort) Then
                        Label2(4).Caption = rs!Ort
                    End If
                Else
                    anzeige "rot", "keine gültige Bankleitzahl", Label2(0)
                End If
            Else
                
                anzeige "rot", "keine gültige Bankleitzahl", Label2(0)
                Label2(3).Caption = ""
                Label2(4).Caption = ""
'                Text1(2).SetFocus
                
            End If
            rs.Close: Set rs = Nothing
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = gcNUM & " " & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    
    Fehlermeldung1
End Sub
