VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL26 
   BackColor       =   &H00C0C000&
   Caption         =   "Bereitstellung für die Zentrale"
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
   Icon            =   "frmWKL26.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command6 
      Height          =   495
      Index           =   3
      Left            =   8040
      TabIndex        =   14
      Top             =   5520
      Width           =   1935
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
      Caption         =   "Protokoll"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.CheckBox Check1 
      Caption         =   "alles markieren"
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   3960
      Width           =   1935
   End
   Begin sevCommand3.Command Command6 
      Height          =   375
      Index           =   2
      Left            =   5160
      TabIndex        =   12
      Top             =   1560
      Width           =   1935
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command6 
      Height          =   495
      Index           =   1
      Left            =   8040
      TabIndex        =   11
      Top             =   4320
      Width           =   1935
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
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   1800
      MultiSelect     =   2  'Erweitert
      TabIndex        =   6
      Top             =   3120
      Width           =   6135
   End
   Begin sevCommand3.Command Command6 
      Height          =   495
      Index           =   0
      Left            =   8040
      TabIndex        =   4
      Top             =   4920
      Width           =   1935
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
      Caption         =   "Senden"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   8520
      MultiSelect     =   2  'Erweitert
      Pattern         =   "F*.lzh"
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin sevCommand3.Command Command6 
      Height          =   495
      Index           =   7
      Left            =   9600
      TabIndex        =   1
      Top             =   7800
      Width           =   1935
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
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   20
      Left            =   4440
      TabIndex        =   15
      ToolTipText     =   "Kalender"
      Top             =   1560
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      ToolTip         =   "Wählen Sie hier das Datum aus."
      ToolTipTitle    =   "Kalender"
      ButtonStyle     =   2
      Caption         =   ""
      Image           =   20
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   21
      Left            =   4440
      TabIndex        =   16
      ToolTipText     =   "Kalender"
      Top             =   2040
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      ToolTip         =   "Wählen Sie hier das Datum aus."
      ToolTipTitle    =   "Kalender"
      ButtonStyle     =   2
      Caption         =   ""
      Image           =   20
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "von :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "bis :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lbl36 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   6240
      Width           =   8175
   End
   Begin VB.Label Label1 
      Caption         =   "Wählen Sie die Datei, die Sie senden wollen"
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
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   2760
      Width           =   6135
   End
   Begin VB.Label lblUeberschrift 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Kassendateien für die Zentrale ausgeben"
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
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11895
   End
End
Attribute VB_Name = "frmWKL26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If Check1.Value = vbChecked Then
        For lcount = 0 To List1.ListCount - 1
            List1.Selected(lcount) = True
        Next lcount
    ElseIf Check1.Value = vbUnchecked Then
        For lcount = 0 To List1.ListCount - 1
            List1.Selected(lcount) = False
        Next lcount
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."
    
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
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 20        ' Kalender
            Text1(2).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            Text1(3).Text = Text1(2).Text
            Text1(3).SetFocus
            
        Case Is = 21        ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            Text1(2).SetFocus
            'fertig
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command6_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0 'senden
            If Kopiere Then 'die Auswahl ins Kassout
                giKissFtpMode = 9 'FTPMODE= 9 , Kombimode Kassendateien holen und schicken
                frmWKL38.Show 1
            End If
            Check1.Value = vbUnchecked
            
        Case Is = 2 'suche
            newseek
            
            
        Case Is = 1 'löschen
            Deldat
            Check1.Value = vbUnchecked

        Case Is = 7   'Schließen
            loeschNEW "fList", gdApp
            Unload frmWKL26
        Case 3
            Dim cPfad As String
            cPfad = gcDBPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
            
            zeigeHilfe "LPROTOK", "FProtokoll.txt", cPfad
        
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub newseek()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount      As Long
    Dim ctmp        As String
    Dim cSatz       As String
    Dim cDatname    As String
    Dim cDatum      As String
    Dim czeit       As String
    Dim sSQL        As String
    Dim rs          As Recordset
    Dim cVon        As String
    Dim cBis        As String
    Dim lVon        As Long
    Dim lBis        As Long
    
    Check1.Value = vbUnchecked

    cVon = Text1(2).Text
    cBis = Text1(3).Text
   
    If IsDate(cVon) Then
         lVon = DateValue(cVon)
    Else
         lVon = 0
    End If
    
    If IsDate(cBis) Then
         lBis = DateValue(cBis)
    Else
         lBis = 0
    End If
  
    loeschNEW "fList", gdApp
    CreateTable "FLIST", gdApp
    
    File1.Path = gcDBPfad & "\abschlus"
    File1.Refresh
    
    For lcount = 0 To File1.ListCount - 1
        cDatname = File1.list(lcount)
        
        cDatum = Mid(FileDateTime(gcDBPfad & "\abschlus\" & cDatname), 1, InStr(1, FileDateTime(gcDBPfad & "\abschlus\" & cDatname), " ") - 1)
        czeit = Right(FileDateTime(gcDBPfad & "\abschlus\" & cDatname), 8)

        sSQL = "Insert into flist (Datname,adate,azeit) values "
        sSQL = sSQL & " ( '" & cDatname & "' "
        sSQL = sSQL & " , '" & cDatum & "' "
        sSQL = sSQL & " , '" & czeit & "' "
        sSQL = sSQL & " ) "

        gdApp.Execute sSQL, dbFailOnError
    Next lcount
    
    sSQL = "Select * from flist "
    
    If lVon <> 0 Then
    sSQL = sSQL & " where ADATE between " & lVon & " and " & lBis
    End If
    
    
    sSQL = sSQL & " order by adate desc "
    
    
    Set rs = gdApp.OpenRecordset(sSQL)
    
    List1.Clear
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            cSatz = ""
            If Not IsNull(rs!Datname) Then
                ctmp = Trim(rs!Datname)
                cSatz = ctmp & Space$(14 - Len(ctmp))
                
                If Not IsNull(rs!Adate) Then
                    ctmp = Format(Trim(rs!Adate), "DD.MM.YY")
                End If
                cSatz = cSatz & ctmp & Space$(3)
                
                If Not IsNull(rs!AZEIT) Then
                    ctmp = Trim(rs!AZEIT)
                End If
                cSatz = cSatz & ctmp
            End If
            List1.AddItem cSatz
            
            rs.MoveNext
        Loop
        anzeigeNew "normal", "Markieren Sie eine oder mehrere Datei/en!", lbl36
    Else
        anzeigeNew "Rot", "Keine Dateien gefunden", lbl36
    End If
    rs.Close: Set rs = Nothing
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "newseek"
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Kopiere() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Kopiere = False
    Dim lcount      As Long
    Dim cLBSatz     As String
    Dim cPfad       As String
    Dim sQuell      As String
    Dim sZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    
    
    cPfad = gcDBPfad 'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    
    If List1.ListIndex = 0 And List1.Selected(0) = False Then
    
        anzeigeNew "Rot", "Bitte eine oder mehrere Dateien auswählen!", lbl36
        List1.SetFocus
        Exit Function
    End If
    
    
    lbl36.Caption = ""
    lbl36.Refresh
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) Then
            cLBSatz = Left(List1.list(lcount), 12)
            schreibeFProtokoll "Datei: " & cLBSatz & " nochmals angefordert"
            sQuell = cPfad & "abschlus\" & cLBSatz
            sZiel = cPfad & "kassout\" & cLBSatz
        
            lRet = CopyFile(sQuell, sZiel, lfail)
            If lRet = 0 Then
                anzeigeNew "Rot", "Konnte " & sQuell & " nicht kopieren!", lbl36
                
            End If
        End If
        
    Next lcount
    
    For lcount = 0 To List1.ListCount - 1
        List1.Selected(lcount) = False
    Next lcount
    
    List1.ListIndex = 0
    
    Kopiere = True
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "kopiere"
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function
Private Sub Deldat()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount      As Long
    Dim cLBSatz     As String
    Dim cPfad       As String
    Dim sSQL        As String
    
    cPfad = gcDBPfad 'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If List1.ListIndex = 0 And List1.Selected(0) = False Then
        anzeigeNew "Rot", "Bitte eine oder mehrere Dateien auswählen!", lbl36
        List1.SetFocus
        Exit Sub
    End If
    
    lbl36.Caption = ""
    lbl36.Refresh
   
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) Then
            cLBSatz = Left(List1.list(lcount), 12)
    
            Kill cPfad & "abschlus\" & cLBSatz
            schreibeFProtokoll "Datei: " & cLBSatz & " gelöscht"
            sSQL = "delete from flist where datname = '" & cLBSatz & "'"
            gdApp.Execute sSQL, dbFailOnError
            
        End If
        
    Next lcount
    
    newseek
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Deldat"
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."

    Fehlermeldung1
    

End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    PositionierenWKL26
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Text1(2).Text = ""
    Text1(3).Text = ""
    
    newseek
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL26()
    On Error GoTo LOKAL_ERROR
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL26"
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Bereitstellung der Daten für die Zentrale ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
