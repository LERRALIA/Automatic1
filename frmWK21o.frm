VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWK21o 
   BackColor       =   &H000000C0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   7080
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10200
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
         Height          =   555
         Index           =   3
         Left            =   3840
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2640
         Width           =   1815
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
         Height          =   555
         Index           =   2
         Left            =   3840
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1440
         Width           =   1815
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
         Height          =   555
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2640
         Width           =   1815
      End
      Begin sevCommand3.Command Command3 
         VBButton        =   1
         ButtonStyle     =   2
         BackColor       =   &H000000C0&
         Caption         =   "Protokoll"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   8
         Top             =   5160
         Visible         =   0   'False
         Width           =   2775
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
         Height          =   555
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1440
         Width           =   1815
      End
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         BackColor       =   &H000000C0&
         Caption         =   "Abbrechen"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   6
         Top             =   6360
         Width           =   2775
      End
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         BackColor       =   &H000000C0&
         Caption         =   "Weiter"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   5
         Top             =   5760
         Width           =   2775
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   11
         Left            =   3720
         TabIndex        =   9
         Top             =   5520
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   10
         Left            =   3720
         TabIndex        =   10
         Top             =   4800
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   ","
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   9
         Left            =   3000
         TabIndex        =   11
         Top             =   5520
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   8
         Left            =   2280
         TabIndex        =   12
         Top             =   5520
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   7
         Left            =   1560
         TabIndex        =   13
         Top             =   5520
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   6
         Left            =   840
         TabIndex        =   14
         Top             =   5520
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   5520
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   4
         Left            =   3000
         TabIndex        =   16
         Top             =   4800
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   3
         Left            =   2280
         TabIndex        =   17
         Top             =   4800
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   2
         Left            =   1560
         TabIndex        =   18
         Top             =   4800
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   1
         Left            =   840
         TabIndex        =   19
         Top             =   4800
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   4800
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "0"
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
         Index           =   5
         Left            =   4560
         TabIndex        =   25
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Dukaten in Stück"
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
         Index           =   4
         Left            =   3840
         TabIndex        =   24
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Dukaten als Wert in Euro"
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
         Index           =   3
         Left            =   3840
         TabIndex        =   23
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Gutscheine in Stück"
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
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Eingabeaufforderung"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   6975
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Bargeld als Wert in Euro"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmWK21o"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR

    Unload frmWK21o
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR

    Dim cdatum      As String
    Dim czeit       As String
    ReDim cZeilen(0 To 9) As String

    If IsNumeric(Text1(0).Text) Then
        schreibeProtoAbschluss "Kassenabschluss beginnt---------------"
        schreibeProtoAbschluss gcUserName & "(" & gcBedienerNr & ") Bargeldeingabe   = " & Text1(0).Text & " " & gcWaehrung
        schreibeProtoAbschluss gcUserName & "(" & gcBedienerNr & ") Dukateneingabe   = " & Text1(2).Text & " " & gcWaehrung
        schreibeProtoAbschluss gcUserName & "(" & gcBedienerNr & ") Gutscheine Stück = " & Text1(1).Text & " Stück"
        schreibeProtoAbschluss gcUserName & "(" & gcBedienerNr & ") Dukaten Stück    = " & Text1(3).Text & " Stück"
        
      
    
    
    
        cdatum = DateValue(Now)
        czeit = TimeValue(Now)
    
        'Drucke den Beleg

        cZeilen(0) = "I S T B E S T A N D"
        cZeilen(1) = "-------------------"
        cZeilen(2) = "Bargeld:      " & Text1(0).Text
        cZeilen(3) = "Dukaten EUR:  " & Text1(2).Text
        cZeilen(4) = "Dukaten Stck: " & Text1(3).Text
        cZeilen(5) = "Gutscheine:   " & Text1(1).Text
        cZeilen(6) = gcUserName
        cZeilen(7) = ""
        cZeilen(8) = "Datum:        " & cdatum
        cZeilen(9) = "Zeit:         " & czeit
        
        DruckeArbeitszeitBelegWK20d cZeilen(), 9
                
        
        frmWKL21.LeseDatenWKL21
        frmWKL21.Show 1
        Unload frmWK21o
    Else
        Unload frmWK21o
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    zeigeHilfeDabapfad "ABPRO", "ABPROTO.txt"
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True:
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lbl6(1)

    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand7_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 11 Then
        Text1(lbl6(5).Caption).Text = ""
        Text1(lbl6(5).Caption).SetFocus
    Else
        If (Index = 10 And lbl6(5).Caption = 1) Or (Index = 10 And lbl6(5).Caption = 3) Then
            Text1(lbl6(5).Caption).SetFocus
            Text1(lbl6(5).Caption).SelStart = Len(Text1(lbl6(5).Caption))
        Else
        
            Text1(lbl6(5).Caption).Text = Text1(lbl6(5).Caption).Text & SSCommand7(Index).Caption
            Text1(lbl6(5).Caption).SetFocus
            Text1(lbl6(5).Caption).SelStart = Len(Text1(lbl6(5).Caption))
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand7_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer, Index As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    
    Select Case Index
    Case 0
        cValid = "1234567890," & Chr$(8)
    Case 1
        cValid = "1234567890" & Chr$(8)
    Case 2
        cValid = "1234567890," & Chr$(8)
    Case 3
        cValid = "1234567890" & Chr$(8)
    End Select
    
    If InStr(cValid, UCase$(Chr$(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    lbl6(5).Caption = Index
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
