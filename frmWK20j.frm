VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWK20j 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "offene Artikelauswahl"
   ClientHeight    =   6810
   ClientLeft      =   150
   ClientTop       =   1485
   ClientWidth     =   9525
   Icon            =   "frmWK20j.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   6810
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   7680
      TabIndex        =   8
      Top             =   4440
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MaxLength       =   4
      TabIndex        =   0
      Top             =   5760
      Width           =   1095
   End
   Begin MSComctlLib.TreeView List3 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4683
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   7680
      TabIndex        =   2
      Top             =   6120
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
   Begin sevCommand3.Command Command1 
      Height          =   735
      Index           =   0
      Left            =   7680
      TabIndex        =   1
      Top             =   5160
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
      Caption         =   "Auswählen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   9255
   End
   Begin VB.Label Label15 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   7455
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "offene Artikelauswahl"
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
      TabIndex        =   3
      Top             =   120
      Width           =   8415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   9360
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmWK20j"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz         As String
    Dim cBedNr          As String
    Dim cBedNrBon       As String
    Dim cSQL            As String
    
    Select Case Index
        Case 0
            cLBSatz = Trim(Text1.Text)
            
            If IsNumeric(cLBSatz) Then
                HoleUnterbrochenenBonWK20j_ARTAUSWAHL cLBSatz
                Unload frmWK20j
            Else
                anzeigeNew "rot", "Bitte einen Eintrag in der Liste auswählen!", Label15
            End If
                
        Case 1
            Unload frmWK20j
        Case 2
            cLBSatz = Trim(Text1.Text)
            
            If IsNumeric(cLBSatz) Then
                cSQL = "Delete from ARTAUSWAHL where LFDNR = " & cLBSatz & " "
                gdBase.Execute cSQL, dbFailOnError
                
                List2.Clear
                zeigvorgaenge
                anzeigeNew "normal", "V-Nr eingeben oder einen Eintrag in der Liste auswählen!", Label15
            Else
                anzeigeNew "rot", "Bitte einen Eintrag in der Liste auswählen!", Label15
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigvorgaenge()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim iFeld As Integer
    Dim cLBSatz As String
    Dim lcount As Long
    
    List1.Clear

    List3.Nodes.Clear
    List1.AddItem "Bed     KundNr KundenName                       Zwischensumme  V-Nr"
    
    cSQL = "Select distinct LFDNR, BEDNR, KDNR, KDNAME, ZSUM, ADATE from ARTAUSWAHL "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!BEDNR) Then
                cFeld = rsrs!BEDNR
            Else
                cFeld = ""
            End If
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!KdNr) Then
                cFeld = rsrs!KdNr
            Else
                cFeld = ""
            End If
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!KdName) Then
                cFeld = rsrs!KdName
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!ZSUM) Then
                cFeld = rsrs!ZSUM
            Else
                cFeld = ""
            End If
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!LFDNR) Then
                iFeld = rsrs!LFDNR
            Else
                iFeld = 0
            End If
            cFeld = Trim$(Str$(iFeld))
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!Adate) Then
                cFeld = Format(rsrs!Adate, "DD.MM.YY")
            Else
                cFeld = ""
            End If

            cFeld = Space$(12 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            List3.Nodes.Add Text:=cLBSatz
            
            List3.Nodes(List3.Nodes.Count).BackColor = vbYellow
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lcount = 1 Then
        List3.Nodes(1).Selected = True
        List3_NodeClick List3.Nodes(1)
        List3.Nodes(1).BackColor = vbBlue
        List3.Nodes(1).ForeColor = vbWhite
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigvorgaenge"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub zeigeParkvorgänge(cNum As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim iFeld As Integer
    Dim cLBSatz As String
    
    List2.Clear

    cSQL = "Select lbtext  from ARTAUSWAHL where lfdnr = " & cNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!lbtext) Then
                cFeld = rsrs!lbtext
            Else
                cFeld = ""
            End If
            
            List2.AddItem cFeld
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigeParkvorgänge"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    
    
    zeigvorgaenge
    anzeigeNew "normal", "V-Nr eingeben oder einen Eintrag in der Liste auswählen!", Label15

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List3_GotFocus()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    
    If List3.SelectedItem Is Nothing Then
        Exit Sub
    Else
        For i = 1 To List3.Nodes.Count
            If List3.Nodes(i).Selected = True Then

                Text1.Text = Mid(List3.Nodes(i), 67, 4)
                Text1.Refresh
                If Text1.Text <> "" Then
                    zeigeParkvorgänge Text1.Text
                End If
                Exit For
            Else

            End If

        Next i
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub List3_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    
    If List3.SelectedItem Is Nothing Then
        Exit Sub
    Else
        For i = 1 To List3.Nodes.Count
            If List3.Nodes(i).Selected = True Then
            
                Text1.Text = Mid(List3.Nodes(i), 67, 4)
                Text1.Refresh
                
                If Text1.Text <> "" Then
                    zeigeParkvorgänge Text1.Text
                End If
                Exit For
            Else

            End If
        Next i
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_NodeClick"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
   
    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    
    cValid = gcNUM
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Beep
    End If
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command1_Click 0
    ElseIf KeyCode = vbKeyEscape Then
        Command1_Click 1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil offene Artikelauswahl ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

