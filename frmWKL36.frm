VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL36 
   BackColor       =   &H00C0C000&
   Caption         =   "Tabellen - Layout bearbeiten"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL36.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8160
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command7 
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4560
      Width           =   2175
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
      Caption         =   "Aktualisieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command6 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   4560
      Width           =   2295
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   4680
      TabIndex        =   7
      Top             =   1680
      Width           =   3255
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   975
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
      Caption         =   "<"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3120
      Width           =   975
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
      Caption         =   "<<"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   2520
      Width           =   975
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
      Caption         =   ">>"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   1920
      Width           =   975
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
      Caption         =   ">"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   4560
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
      Caption         =   "Schließe&n"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reihenfolge und Anzeige der Tabellenspalten bestimmen Sie hier."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4680
      TabIndex        =   10
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "verfügbare Tabellenspalten"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabellenbearbeitung"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   7920
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmWKL36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTab As String
Private Sub cmdGo_Click()
    On Error GoTo LOKAL_ERROR
    
    Layoutauflisten
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdGo_Click"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
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
Private Sub Layoutauflisten()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim sTemp As String
    Dim i As Byte
    
    List1.Clear
    List2.Clear
    
    If UCase(sTab) = "ZBON" Then
        sSQL = " Select * from ZBONlay where tabname = '" & sTab & "'"
        sSQL = sSQL & " order by Reihenf "
    Else
    
        sSQL = " Select * from Tablay" & srechnertab & " where tabname = '" & sTab & "'"
        sSQL = sSQL & " order by Reihenf "
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If rsrs.EOF Then
    
    Else
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Spaltenna) Then
                sTemp = Trim(rsrs!Spaltenna)
            Else
                sTemp = ""
            End If
            List1.AddItem sTemp
            rsrs.MoveNext
        Loop
    End If
        
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Layoutauflisten"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim bFound As Boolean
    Dim lcount As Long
    
    bFound = False
    
    If List1.ListCount = 0 Then
        Exit Sub
    End If
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
                bFound = True
        End If
    Next lcount
    
    If bFound Then
        
        For lcount = 0 To List1.ListCount - 1
            If List1.Selected(lcount) Then
                List2.AddItem List1.list(lcount)
                List1.RemoveItem lcount
                Exit For
            End If
        Next lcount
    
    Else
        List2.AddItem List1.list(List1.TopIndex)
        List1.RemoveItem List1.TopIndex
    End If
    
    List1.Refresh
    List2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long

    If List1.ListCount = 0 Then
        Exit Sub
    End If

    For lcount = 0 To List1.ListCount - 1
        List2.AddItem List1.list(lcount)
    Next lcount
        
    List1.Clear
    List2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long

    If List2.ListCount = 0 Then
        Exit Sub
    End If

    For lcount = 0 To List2.ListCount - 1
        List1.AddItem List2.list(lcount)
    Next lcount
        
    List2.Clear
    List1.Refresh
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim bFound As Boolean
    Dim lcount As Long
    
    If List2.ListCount = 0 Then
        Exit Sub
    End If
    
    
    bFound = False
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    
    If bFound Then
        
        For lcount = 0 To List2.ListCount - 1
            If List2.Selected(lcount) Then
                List1.AddItem List2.list(lcount)
                List2.RemoveItem lcount
                Exit For
            End If
        Next lcount
    
    Else
    
        List1.AddItem List2.list(List2.TopIndex)
        List2.RemoveItem List2.TopIndex
    End If
    
    List1.Refresh
    List2.Refresh
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command5_Click()
    On Error GoTo LOKAL_ERROR

    Unload frmWKL36
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click()
    On Error GoTo LOKAL_ERROR
    
    Speichertablay srechnertab
    
    Command5_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Speichertablay(srechnertab As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lcount As Long
    Dim sSpalte As String
    Dim stabnameE As String
    
    List2.Refresh
    
    If List2.ListCount = 0 Then
        Exit Sub
    End If
    
    If UCase(sTab) = "ZBON" Then
        stabnameE = "ZBONLAY"
    Else
        stabnameE = "TABLAY" & srechnertab
    End If
    
    sSQL = "Update " & stabnameE & " set REIHENF = '99' "
    sSQL = sSQL & " where tabname = '" & sTab & "' "
    SQL_Befehl_ausführen sSQL
   
    
    For lcount = 0 To List2.ListCount - 1

        sSpalte = Trim(List2.list(lcount))
        sSQL = "Update " & stabnameE & " set REIHENF = " & lcount
        sSQL = sSQL & " where tabname = '" & Trim(sTab) & "'"
        sSQL = sSQL & " and spaltenna = '" & sSpalte & "'"
        SQL_Befehl_ausführen sSQL
    Next lcount

    If UCase(sTab) = "BEAKU" Or UCase(sTab) = "BEALIEF" Or UCase(sTab) = "MASTEMP" _
    Or UCase(sTab) = "KOPFMAIL" Or UCase(sTab) = "STADAPRI" Or UCase(sTab) = "STADAPRIB" Or UCase(sTab) = "BESTELLUNG" Then
        sSQL = "Update " & stabnameE & " set ANZEIGE = 'J' "
        sSQL = sSQL & " where tabname = '" & sTab & "' "
        sSQL = sSQL & " and REIHENF = 99 "
        SQL_Befehl_ausführen sSQL
    Else
        sSQL = "Update " & stabnameE & " set ANZEIGE = 'N' "
        sSQL = sSQL & " where tabname = '" & sTab & "' "
        sSQL = sSQL & " and REIHENF = 99 "
        SQL_Befehl_ausführen sSQL
    End If
    sSQL = "Update " & stabnameE & " set ANZEIGE = 'J' "
    sSQL = sSQL & " where tabname = '" & sTab & "' "
    sSQL = sSQL & " and REIHENF <> 99 "
    SQL_Befehl_ausführen sSQL
    
    sSQL = "Update " & stabnameE & " set ANZEIGE = 'J' "
    sSQL = sSQL & " where tabname = '" & sTab & "' "
    sSQL = sSQL & " and spaltenbez = '" & gsZSpalte & "'"
    SQL_Befehl_ausführen sSQL
    
    sSQL = "Update " & stabnameE & " set ANZEIGE = 'J' "
    sSQL = sSQL & " where tabname = '" & sTab & "' "
    sSQL = sSQL & " and spaltenbez = '" & gsZSpalte1 & "'"
    SQL_Befehl_ausführen sSQL
    
    sSQL = "Update " & stabnameE & " set ANZEIGE = 'J' "
    sSQL = sSQL & " where tabname = '" & sTab & "' "
    sSQL = sSQL & " and spaltenbez = '" & gsZSpalte2 & "'"
    SQL_Befehl_ausführen sSQL
    
    sSQL = "Update " & stabnameE & " set ANZEIGE = 'J' "
    sSQL = sSQL & " where tabname = '" & sTab & "' "
    sSQL = sSQL & " and spaltenbez = '" & gsZSpalte3 & "'"
    SQL_Befehl_ausführen sSQL
    
    sSQL = "Update " & stabnameE & " set ANZEIGE = 'J' "
    sSQL = sSQL & " where tabname = '" & sTab & "' "
    sSQL = sSQL & " and spaltenbez = '" & gsZSpalte4 & "'"
    SQL_Befehl_ausführen sSQL
    
    sSQL = "Update " & stabnameE & " set ANZEIGE = 'J' "
    sSQL = sSQL & " where tabname = '" & sTab & "' "
    sSQL = sSQL & " and spaltenbez = '" & gsZSpalte5 & "'"
    SQL_Befehl_ausführen sSQL
    
    List2.Clear
    cmdGo_Click
    Command5_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Speichertablay"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub Command7_Click()
    On Error GoTo LOKAL_ERROR
    
    
    delete sTab
    Tabcheck sTab 'Tabellencheck
    Layoutauflisten
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub delete(sTab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If UCase(sTab) = "ZBON" Then
        sSQL = " Delete from ZBONlay where tabname = '" & sTab & "'"
        SQL_Befehl_ausführen sSQL
        
    Else
        sSQL = " Delete from Tablay" & srechnertab & " where tabname = '" & sTab & "'"
        SQL_Befehl_ausführen sSQL
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "delete"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    PositionierenWKL36
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift

    sTab = gstab
    Tabcheck sTab 'Tabellencheck
    Layoutauflisten
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL36()
    On Error GoTo LOKAL_ERROR

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL36"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Command1_Click
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2_dblClick()
    On Error GoTo LOKAL_ERROR
    
    Command4_Click
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

