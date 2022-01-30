VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form dlgKopieren 
   Caption         =   " - Dateien kopieren"
   ClientHeight    =   7890
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   8670
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   3255
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "alle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin sevCommand3.Command Command0 
         Height          =   420
         Index           =   0
         Left            =   1320
         TabIndex        =   21
         ToolTipText     =   "Kalender"
         Top             =   120
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   741
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Filiale:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Datum:"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   19
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Kiste:"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3360
      Pattern         =   "*.lzh"
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdMehrfachZurück 
      Caption         =   "<<"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton cmdEinfachZurück 
      Caption         =   "<"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmdMehrfachauswahl 
      Caption         =   ">>"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdEinfachAuswahl 
      Caption         =   ">"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Schließen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   7320
      Width           =   3015
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7320
      Width           =   3015
   End
   Begin VB.ListBox List2 
      Height          =   4350
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "von:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6960
      Width           =   8415
   End
   Begin VB.Label Label3 
      Caption         =   "nach:"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "von:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Sicherungsdateien kopieren"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "dlgKopieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelButton_Click()
On Error GoTo LOKAL_ERROR

    Unload Me
 
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CancelButton_Click"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub cmdEinfachAuswahl_Click()
On Error GoTo LOKAL_ERROR

    If List1.ListIndex >= 0 Then
        List2.AddItem List1.list(List1.ListIndex)
        List1.RemoveItem List1.ListIndex
        
        anzeige "normal", "", Label5
        If List2.ListCount > 0 Then
            OKButton.Caption = "Senden"
        Else
            OKButton.Caption = "OK"
        End If
    Else
        anzeige "rot", "Wählen Sie bitte eine Datei aus!", Label5
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEinfachAuswahl_Click"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub cmdEinfachZurück_Click()
On Error GoTo LOKAL_ERROR

    If List2.ListIndex >= 0 Then
        List1.AddItem List2.list(List2.ListIndex)
        List2.RemoveItem List2.ListIndex
        
        anzeige "normal", "", Label5
        If List2.ListCount > 0 Then
            OKButton.Caption = "Senden"
        Else
            OKButton.Caption = "OK"
        End If
    Else
        anzeige "rot", "Wählen Sie bitte eine Datei aus!", Label5
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEinfachZurück_Click"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub cmdMehrfachZurück_Click()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim iAnzFil As Integer
    
    iAnzFil = List2.ListCount - 1
    For i = 0 To iAnzFil
        List1.AddItem List2.list(i)
    Next i
    
    List2.Clear
    
    If List2.ListCount > 0 Then
        OKButton.Caption = "Senden"
    Else
        OKButton.Caption = "OK"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdMehrfachZurück_Click"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub cmdMehrfachauswahl_Click()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim iAnzFil As Integer
    iAnzFil = List1.ListCount - 1
    For i = 0 To iAnzFil
        List2.AddItem List1.list(i)
    Next i
    List1.Clear
    
    If List2.ListCount > 0 Then
        OKButton.Caption = "Senden"
    Else
        OKButton.Caption = "OK"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdMehrfachauswahl_Click"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
'Private Sub cmdProto_Click()
'On Error GoTo LOKAL_ERROR
'
'    Screen.MousePointer = 11
'    zeigeHilfeDabapfad "LPROTOK", "Sicherungsdat.txt"
'    Screen.MousePointer = 0
'
'    Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "cmdProto_Click"
'    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."
'    Fehlermeldung1
'End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(1).Text = Format(Datumschreiben11a(3500, 3500), "DD.MM.YY")
    Text1_KeyUp 1, vbKeyReturn, 0
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim sFiletime As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    List1.Clear

    File1.Pattern = "N*.mdb"
    File1.Path = cPfad & "WVOUTSIC"
    For i = 0 To File1.ListCount - 1
        sFiletime = FileDateTime(cPfad & "WVOUTSIC\" & File1.list(i))
        List1.AddItem File1.list(i) & "     " & sFiletime
    Next i
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sFiletime As String
    Dim cPfad As String
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Label1
    

    OKButton.Caption = "Auswählen"
    OKButton.Refresh
    
    altesloeschen "WVOUTSIC"

    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    File1.Pattern = "N*.mdb"
    File1.Path = cPfad & "WVOUTSIC"
    Label2.Caption = Label2.Caption & " " & cPfad & "WVOUTSIC"
    Label3.Caption = Label3.Caption & " " & cPfad & "WVOUT"
    For i = 0 To File1.ListCount - 1
        sFiletime = Format(FileDateTime(cPfad & "WVOUTSIC\" & File1.list(i)), "DD.MM.YY") & " " & Format(FileDateTime(cPfad & "WVOUTSIC\" & File1.list(i)), "MM:HH:SS")
        List1.AddItem File1.list(i) & "     " & sFiletime
    Next i
    
    anzeige "normal", "Wählen Sie bitte eine Datei aus!", Label5
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub altesloeschen(sUnterver As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnz        As Long
    Dim lcount      As Long
    Dim lHeute      As Long
    Dim lDateiDatum As Long
    Dim cdatei      As String
    Dim cPfad       As String
    
    lHeute = Fix(Now)
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & sUnterver & "\"
    
    File1.Path = cPfad
    File1.Pattern = "*.*"
    File1.Refresh
    
    lAnz = File1.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = File1.list(lcount)
        lDateiDatum = FileDateTime(cPfad & cdatei)
        If lHeute - lDateiDatum > 90 Then

            Kill cPfad & cdatei
            
        End If
    Next lcount
    
    File1.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "altesloeschen"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Label4_Click()

End Sub

Private Sub OKButton_Click()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Select Case OKButton.Caption
        Case "OK"
            CancelButton_Click
        Case "Senden"
            For i = 0 To List2.ListCount - 1
                CopyFile cPfad & "WVOUTSIC\" & Mid$(List2.list(i), 1, InStr(1, List2.list(i), " ")), cPfad & "WVOUT\" & Mid$(List2.list(i), 1, InStr(1, List2.list(i), " ")), False
            Next i

            giKissFtpMode = 16
            frmWKL38.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "OKButton_Click"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim cKiste As String

    List3.Clear
    
    If KeyCode = vbKeyReturn Then
        Select Case Index
            
            Case 0 'fil
                For i = 0 To List1.ListCount - 1
                    If Val(Mid(List1.list(i), 2, 2)) = Val(Text1(Index).Text) Then
                        List3.AddItem List1.list(i)
                    
                    End If
                Next i
                
                List1.Clear
                
                For i = 0 To List3.ListCount - 1
                    List1.AddItem List3.list(i)
                Next i
            
            Case 1 'Datum
            
                For i = 0 To List1.ListCount - 1
                    If Mid(List1.list(i), Len(List1.list(i)) - 16, 8) = Text1(Index).Text Then
                        List3.AddItem List1.list(i)
                    End If
                Next i
                
                List1.Clear
                
                For i = 0 To List3.ListCount - 1
                    List1.AddItem List3.list(i)
                Next i
            
            Case 2 'Kiste
            
                For i = 0 To List1.ListCount - 1
                    If Mid(List1.list(i), 6, InStr(1, List1.list(i), ".") - 6) = Trim(Text1(Index).Text) Then
                        List3.AddItem List1.list(i)
                    End If
                Next i
                
                List1.Clear
                
                For i = 0 To List3.ListCount - 1
                    List1.AddItem List3.list(i)
                Next i
            
        End Select
    End If
    
    If List1.ListCount = 0 Then
        Command1_Click
    End If

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Sicherungsdateien kopieren ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
