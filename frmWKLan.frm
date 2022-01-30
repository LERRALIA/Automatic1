VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKLan 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Änderung von Artikeldaten"
   ClientHeight    =   6645
   ClientLeft      =   2670
   ClientTop       =   1575
   ClientWidth     =   6105
   Icon            =   "frmWKLan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   6645
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame3 
      Caption         =   "Aktion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   20
      Top             =   5520
      Width           =   6015
      Begin Threed.SSCommand SSCommand1 
         Height          =   735
         Index           =   1
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "Schließen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "Speichern"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Daten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   6015
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1920
         MaxLength       =   13
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1920
         MaxLength       =   13
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1920
         MaxLength       =   13
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Bonusfähig:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Preisschutz:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Art. im Geschäft:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Rabattfähig:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "3.EAN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "2.EAN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "EAN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ausgewählter Artikel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   26
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Lieferantenname:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   1080
         TabIndex        =   24
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Lief.Nr.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Bezeichnung:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Art.Nr.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmWKLan"
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
Private Sub HoleDatenWKLan()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lArtNr As Long
    
    lArtNr = Val(Label2(0).Caption)
    
    sSQL = "Select * from ARTIKEL where artnr = " & lArtNr
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!EAN) Then
            Text1(0).Text = rsrs!EAN
        Else
            Text1(0).Text = ""
        End If
        If Not IsNull(rsrs!EAN2) Then
            Text1(1).Text = rsrs!EAN2
        Else
            Text1(1).Text = ""
        End If
        If Not IsNull(rsrs!EAN3) Then
            Text1(2).Text = rsrs!EAN3
        Else
            Text1(2).Text = ""
        End If
        If Not IsNull(rsrs!RABATT_OK) Then
            Text1(3).Text = rsrs!RABATT_OK
        Else
            Text1(3).Text = "J"
        End If
        If Not IsNull(rsrs!GEFUEHRT) Then
            Text1(4).Text = rsrs!GEFUEHRT
        Else
            Text1(4).Text = "J"
        End If
        If Not IsNull(rsrs!PREISSCHU) Then
            Text1(5).Text = rsrs!PREISSCHU
        Else
            Text1(5).Text = "N"
        End If
        If Not IsNull(rsrs!BONUS_OK) Then
            Text1(6).Text = rsrs!BONUS_OK
        Else
            Text1(6).Text = "J"
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleDatenWKLan"
    Fehler.gsFehlertext = "Im Programmteil Artikeländerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub SchreibeDatenWKLan()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs        As Recordset
    Dim lArtNr      As Long
    Dim cFeld       As String
    Dim lcount      As Long
    Dim lRet        As Long
    Dim sSQL        As String
    Dim bEAN        As Boolean
    Dim bEAN2       As Boolean
    Dim bEAN3       As Boolean
    Dim cTmp As String
    
    bEAN = False
    bEAN2 = False
    bEAN3 = False
    
    For lcount = 0 To 2
        cFeld = Text1(lcount).Text
        cFeld = Trim$(cFeld)
        If cFeld <> "" Then
            lRet = fnPruefeEANWert(cFeld)
            If lRet <> 0 Then
                MsgBox "Der eingegebene Wert ist kein gültige EAN-Code!", vbCritical, "STOP!"
                Text1(lcount).SetFocus
                Exit Sub
            End If
        End If
    Next lcount
    
    For lcount = 3 To 6
        cFeld = Text1(lcount).Text
        cFeld = Trim$(cFeld)
        If cFeld <> "J" And cFeld <> "N" Then
            MsgBox "Bitt geben Sie einen gültigen Wert ein ( J / N ) !", vbCritical, "STOP!"
            Text1(lcount).SetFocus
            Exit Sub
        End If
    Next lcount
    
    
    
    lArtNr = Val(Label2(0).Caption)
    sSQL = "Select * from ARTIKEL where artnr = " & lArtNr
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
        
        'Hat sich der EAN(1) geändert
    
        If Not IsNull(rsrs!EAN) Then
            If Trim(CStr(rsrs!EAN)) <> Trim$(Text1(0).Text) Then
                bEAN = True
            End If
        Else
            bEAN = True
        End If
        
        'Hat sich der EAN2 geändert
        
        If Not IsNull(rsrs!EAN2) Then
            If Trim(CStr(rsrs!EAN2)) <> Trim$(Text1(1).Text) Then
                bEAN2 = True
            End If
        Else
            bEAN2 = True
        End If
        
        'Hat sich der EAN3 geändert
        
        If Not IsNull(rsrs!EAN3) Then
            If Trim(CStr(rsrs!EAN3)) <> Trim$(Text1(2).Text) Then
                bEAN3 = True
            End If
        Else
            bEAN3 = True
        End If
        
        rsrs!RABATT_OK = Trim$(Text1(3).Text)
        rsrs!GEFUEHRT = Trim$(Text1(4).Text)
        rsrs!PREISSCHU = Trim$(Text1(5).Text)
        rsrs!BONUS_OK = Trim$(Text1(6).Text)
        rsrs!LASTDATE = DateValue(Now)
        rsrs!LASTTIME = TimeValue(Now)
        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    If bEAN Then
        cTmp = Trim$(Text1(0).Text)
        Artikelveraenderung CStr(lArtNr), cTmp, "WE aus E ÄNDERN", "EAN"
    End If
    
    If bEAN2 Then
        cTmp = Trim$(Text1(1).Text)
        Artikelveraenderung CStr(lArtNr), cTmp, "WE aus E ÄNDERN", "EAN2"
    End If
    
    If bEAN3 Then
        cTmp = Trim$(Text1(2).Text)
        Artikelveraenderung CStr(lArtNr), cTmp, "WE aus E ÄNDERN", "EAN3"
    End If
    
    Unload frmWKLan
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKLan"
    Fehler.gsFehlertext = "Im Programmteil Artikeländerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    For lcount = 0 To 6
        Text1(lcount).Text = ""
    Next lcount
    
    Label2(0).Caption = frmWKL15!Label2(2).Caption
    Label2(1).Caption = frmWKL15!Label2(0).Caption
    Label2(2).Caption = frmWKL15!Text1(4).Text
    Label2(3).Caption = frmWKL15!Label2(4).Caption
    
    HoleDatenWKLan
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikeländerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Select Case Index
        Case Is = 0
        
            If checkthisean(Trim$(Text1(0).Text), Trim$(Label2(0).Caption)) = True Then

            Else
                Text1(0).SetFocus
                Screen.MousePointer = 0
                MsgBox "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden.", vbInformation, "Winkiss Hinweis:"
                
                Exit Sub
            End If
            
            If checkthisean(Trim$(Text1(1).Text), Trim$(Label2(0).Caption)) = True Then

            Else
                Text1(1).SetFocus
                Screen.MousePointer = 0
                MsgBox "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden.", vbInformation, "Winkiss Hinweis:"
                
                Exit Sub
            End If
            
            If checkthisean(Trim$(Text1(2).Text), Trim$(Label2(0).Caption)) = True Then

            Else
                Text1(2).SetFocus
                Screen.MousePointer = 0
                MsgBox "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden.", vbInformation, "Winkiss Hinweis:"
                
                Exit Sub
            End If
            SchreibeDatenWKLan
            
        Case Is = 1
            Unload frmWKLan
            
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikeländerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikeländerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case 0 To 2
            cValid = "1234567890" & Chr$(8)
        Case 3 To 6
            cValid = "JN" & Chr$(8)
    End Select
    If cZeichen <> Chr$(8) Then
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Artikeländerung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Artikeländerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


