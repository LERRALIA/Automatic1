VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL73 
   Caption         =   "Kassenprotokolle"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL73.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmWKL73.frx":0442
      Left            =   7800
      List            =   "frmWKL73.frx":0449
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin sevCommand3.Command Command1 
         Height          =   285
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Width           =   615
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "alle"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "filtern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   0
         Left            =   720
         TabIndex        =   16
         ToolTipText     =   "Kalender"
         Top             =   0
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
      Begin VB.Label Label6 
         Caption         =   "Datum:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   4
      Left            =   9600
      TabIndex        =   9
      Top             =   2280
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
      Caption         =   "Ein/Aus Zhl"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   9600
      TabIndex        =   8
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
      Caption         =   "Tagesartikel"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   6
      Top             =   1680
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
   End
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   8760
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   3
      Top             =   1080
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
      Caption         =   "Ansehen"
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   5
      Left            =   9600
      TabIndex        =   17
      Top             =   2880
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
      Caption         =   "Bargeld"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Die ersten 4 Ziffern stehen für das Datum. Die letzten 4 Ziffern stehen für die Uhrzeit (Stunde und Minute)."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   0
      Left            =   4440
      TabIndex        =   7
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kassenprotokolle"
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
      TabIndex        =   2
      Top             =   120
      Width           =   9135
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
      TabIndex        =   1
      Top             =   7920
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL73"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(1).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
    'fertig
    Text1_KeyUp 1, vbKeyReturn, 0
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub


Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR
    
    fuelleliste
    Text1(1).Text = ""
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR
    
    If Text1(1).Text <> "" Then
        Text1_KeyUp 1, vbKeyReturn, 0
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim cKiste As String
    Dim cStueckdatum As String
    
    List3.Clear
    
    If KeyCode = vbKeyReturn Then
        Select Case Index
            
            Case 1 'Datum
                cStueckdatum = Left(Text1(Index).Text, 5)
                'neu 09.01.14
                cStueckdatum = Format(cStueckdatum, "MM.DD")
                'ende neu
                cStueckdatum = SwapStr(cStueckdatum, ".", "")
                

            
                For i = 0 To List1.ListCount - 1
                    If Left(List1.list(i), 4) = cStueckdatum Then
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
        fuelleliste
    End If
    
    If List1.ListCount > 1 Then
        anzeige "normal", List1.ListCount & " Dateien stehen zur Verfügung", Label1(4)
    ElseIf List1.ListCount = 1 Then
        anzeige "normal", "1 Datei steht zur Verfügung", Label1(4)
    ElseIf List1.ListCount = 0 Then
        anzeige "normal", "Es steht keine Datei zur Verfügung", Label1(4)
    End If
            

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "ABPRO"
    
    Dim cPfad1 As String
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    cPfad1 = cPfad1 & "ABPROSIC"
    
    Dim lRet As Long
    Dim lfail As Long

    Select Case Index
        Case 0
            Unload frmWKL73
        Case 1 'ansehen
            If List1.ListIndex >= 0 Then
                anzeige "normal", "", Label1(4)
                Screen.MousePointer = 11
                zeigeHilfeDabapfad "ABPRO", List1.list(List1.ListIndex)
                Screen.MousePointer = 0
            Else
                anzeige "rot", "Wählen Sie bitte eine Datei aus!", Label1(4)
            End If
        Case 2 'löschen /copy
            If List1.ListIndex >= 0 Then
                anzeige "normal", "", Label1(4)
                
                lRet = CopyFile(cPfad & "\" & List1.list(List1.ListIndex), cPfad1 & "\" & List1.list(List1.ListIndex), lfail)
                Kill cPfad & "\" & List1.list(List1.ListIndex)
                fuelleliste
            Else
                anzeige "rot", "Wählen Sie bitte eine Datei aus!", Label1(4)
            End If
        Case 3 '
            tagesArtikelkum
        Case 4 '
'            Ein aus Zahlung
            frmWKL119.Show 1
        Case 5
            zeigeHilfeDabapfad "LPROTOK", "Bargeld_Handling.txt"
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command5_Click"
        Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub tagesArtikelkum()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Screen.MousePointer = 11
    
    loeschNEW "afcd03", gdBase
    CreateTable "AFCD03", gdBase
    
    sSQL = "Insert into AFCD03 Select a.LINR "
    sSQL = sSQL & ", a.AArtnr"
    sSQL = sSQL & ", a.ABEZEICH"
    sSQL = sSQL & ", a.AMenge as ameng"
    sSQL = sSQL & ", a.APreis as aprei"
    sSQL = sSQL & ", a.ALEKPR "
    sSQL = sSQL & ", a.AMWSK as AMWST"
    sSQL = sSQL & ", 0 as Bestand "
    sSQL = sSQL & " "
    sSQL = sSQL & " from afcbuch a  "
    sSQL = sSQL & " where aflag = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update AFCD03 "
    sSQL = sSQL & " set ne = (aPrei * 100) /(100 + " & gdMWStV & ") - (ALEKPR * AMeng)"
    sSQL = sSQL & " where amwst ='V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update AFCD03 "
    sSQL = sSQL & " set ne = (aPrei * 100) /(100 + " & gdMWStO & ") - (ALEKPR * AMeng)"
    sSQL = sSQL & " where amwst ='O' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update AFCD03 "
    sSQL = sSQL & " set ne = (aPrei * 100) /(100 + " & gdMWStE & ") - (ALEKPR * AMeng)"
    sSQL = sSQL & " where amwst ='E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update AFCD03 "
    sSQL = sSQL & " set nettopreis = (aPrei * 100) /(100 + " & gdMWStV & ") "
    sSQL = sSQL & " where amwst ='V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update AFCD03 "
    sSQL = sSQL & " set nettopreis = (aPrei * 100) /(100 + " & gdMWStO & ") "
    sSQL = sSQL & " where amwst ='O' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update AFCD03 "
    sSQL = sSQL & " set nettopreis = (aPrei * 100) /(100 + " & gdMWStE & ") "
    sSQL = sSQL & " where amwst ='E' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    loeschNEW "afcd3", gdBase
    CreateTable "AFCD3", gdBase
    
    sSQL = "Insert into AFCD3 Select a.LINR "
    sSQL = sSQL & ", a.AArtnr"
    sSQL = sSQL & ", a.ABEZEICH"
    sSQL = sSQL & ", sum(a.AMENG)as AMenge"
    sSQL = sSQL & ", sum(a.APREI)as APreis"
    sSQL = sSQL & ", sum(a.ne)as nse"
    sSQL = sSQL & ", sum(a.nettopreis)as nettopr"
    sSQL = sSQL & ", a.ALEKPR "
    sSQL = sSQL & " "
    sSQL = sSQL & " from AFCD03 a "
    sSQL = sSQL & " group by a.AArtnr"
    sSQL = sSQL & ", a.LINR "
    sSQL = sSQL & ", a.ABEZEICH"
    sSQL = sSQL & ", a.ALEKPR "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Update afcbuch  "
'    sSQL = sSQL & " set aflag = 1 where aflag = 0 "
'    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update AFCD3 inner join Artikel on AFCD3.AARTNR = Artikel.Artnr "
    sSQL = sSQL & " set AFCD3.BESTAND = ARTIKEL.BESTAND "
    sSQL = sSQL & " , AFCD3.FARBNR = val(ARTIKEL.awm) "
    gdBase.Execute sSQL, dbFailOnError
    
    BringFarbeInsSpiel "AFCD3", gdBase
    
    If gbKUMSUM = True Then
        'anzeigen
        KUMSUM 1
    Else
        'nicht anzeigen
        KUMSUM 2
    End If
    
    If Datendrin("AFCD3", gdBase) Then
        reportbildschirm "", "aWKL21e"
    Else
        anzeige "rot", "Es sind keine Daten vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "tagesArtikelkum"
    Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    
    fuelleliste
    

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    
'    Anzeige "normal", "", Label1(4)
    
    

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command5_Click"
        Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub fuelleliste()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp    As String
    Dim lcount  As Long
    Dim cExt    As String
    Dim cPfad   As String
    
    cPfad = gcDBPfad
    
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "ABPRO"
    
    File2.Path = cPfad
    File2.Pattern = "*.txt"
    File2.Refresh
    
    List1.Clear
    For lcount = 0 To File2.ListCount - 1
        ctmp = File2.list(lcount)
        ctmp = Trim$(ctmp)
        ctmp = UCase$(ctmp)
        
        If ctmp = "ABPROTO.TXT" Then
        
        Else
            cExt = Right$(ctmp, 3)
            If UCase$(cExt) = "TXT" Then
                List1.AddItem ctmp
                
            End If
        End If
    Next lcount
    If List1.ListCount > 1 Then
        anzeige "normal", List1.ListCount & " Dateien stehen zur Verfügung", Label1(4)
    ElseIf List1.ListCount = 1 Then
        anzeige "normal", "1 Datei steht zur Verfügung", Label1(4)
    ElseIf List1.ListCount = 0 Then
        anzeige "normal", "Es steht keine Datei zur Verfügung", Label1(4)
    End If
    
    List1.Refresh
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelleliste"
    Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."
    Fehlermeldung1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW "afcd03", gdBase
    loeschNEW "afcd3", gdBase
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
Private Sub List1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Command5_Click 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_DblClick"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
