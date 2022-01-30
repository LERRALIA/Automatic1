VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL201 
   Caption         =   "Artikel wiederherstellen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command3 
      Height          =   405
      Index           =   1
      Left            =   7400
      TabIndex        =   6
      Top             =   8010
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Wiederherstellen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame5"
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   11775
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5820
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   2
         Top             =   600
         Width           =   11415
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11415
      End
   End
   Begin sevCommand3.Command Command3 
      Height          =   405
      Index           =   0
      Left            =   9480
      TabIndex        =   0
      Top             =   8010
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin sevCommand3.Command Command3 
      Height          =   405
      Index           =   2
      Left            =   5310
      TabIndex        =   7
      Top             =   8010
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin VB.Label lblanzeige 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   7800
      Width           =   3255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikel wiederherstellen"
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
      Width           =   9495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    With Frame5
        .Height = 6615
        .Left = 0
        .Top = 840
        .Width = 11775
        .BorderStyle = 0
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Artikel wiederherstellen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            Unload frmWKL201
        Case 1
            wiederherstellen
            FuelleListeArtikelWKL201
        Case 2
            zeigeHilfeDabapfad "LPROTOK", "geloeschteArtikel.txt"
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel wiederherstellen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub wiederherstellen()
On Error GoTo LOKAL_ERROR

    Dim bFound As Boolean
    Dim lcount As Long
    Dim iRet As Integer
    Dim cArtNr As String
    Dim cFilnr As String
    Dim cLBSatz As String
    Dim cSQL As String
    

    bFound = False
    
    Screen.MousePointer = 11
            
    For lcount = 0 To List3.ListCount - 1
        If List3.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    
    If bFound Then
        iRet = MsgBox("Wollen Sie nur die markierten Artikel wiederherstellen?", vbYesNoCancel + vbQuestion, "Winkiss Frage:")
        If iRet = vbCancel Then
            Screen.MousePointer = 0
            Exit Sub
        ElseIf iRet = vbYes Then
            
            For lcount = 0 To List3.ListCount - 1
                If List3.Selected(lcount) = True Then
                    cLBSatz = List3.list(lcount)
                    cArtNr = Mid$(cLBSatz, 14, 6)
                    cArtNr = Trim$(cArtNr)
                    
                    ZurückSicherInArtikel cArtNr
                End If
            Next lcount
        Else
            iRet = MsgBox("Wollen Sie wirklich alle Artikel wiederherstellen?", vbYesNoCancel + vbQuestion, "Winkiss Frage:")
            If iRet = vbCancel Then
                Screen.MousePointer = 0
                Exit Sub
            ElseIf iRet = vbYes Then
                ZurückSicherInArtikelALL
            End If
        End If
    Else
        iRet = MsgBox("Wollen Sie wirklich alle Artikel wiederherstellen?", vbYesNoCancel + vbQuestion, "Winkiss Frage:")
        If iRet = vbCancel Then
            Screen.MousePointer = 0
            Exit Sub
        ElseIf iRet = vbYes Then
            ZurückSicherInArtikelALL
        End If
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "wiederherstellen"
    Fehler.gsFehlertext = "Im Programmteil Artikel wiederherstellen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Positionieren

    Schrift Me
    alternativFarbform Me, lblUeberschrift
    LogtoStart Me
    
    If NewTableSuchenDBKombi("Artikelsic", gdBase) = False Then
        CreateArtikelsic
    End If
    
    sSQL = "Delete from Artikelsic where DELDATE < datevalue(now) - 365"
    gdBase.Execute sSQL, dbFailOnError
    
    FuelleListeArtikelWKL201
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikel wiederherstellen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleListeArtikelWKL201()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim cFeld As String
    Dim dWert As Double
    Dim iFil As Integer
    Dim bAnd As Boolean
    
    bAnd = False
    
    List1.Clear
    List3.Clear
    List3.Visible = False
    
    
    List1.AddItem "Datum  Zeit  ArtNr. Artikelbezeichnung                     VK-Preis"
    
    cSQL = "Select * from Artikelsic "
    cSQL = cSQL & "order by  DELDATE desc,DELTIME desc"
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!DELDATE) Then
                cFeld = Format(rsrs!DELDATE, "DD.MM.")
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space(6 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
        
            If Not IsNull(rsrs!DELTIME) Then
                cFeld = Format(rsrs!DELTIME, "HH:MM")
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
        
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!vkpr) Then
                dWert = rsrs!vkpr
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "###,##0.00")
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
'            If Not IsNull(rsrs!ANZAHL) Then
'                dWert = rsrs!ANZAHL
'            Else
'                dWert = 0
'            End If
'            cFeld = Format$(dWert, "#####0")
'            cFeld = Space$(14 - Len(cFeld)) & cFeld
'            cLBSatz = cLBSatz & cFeld & " "
            
'            If Not IsNull(rsrs!filnr) Then
'                dWert = rsrs!filnr
'            Else
'                dWert = gcFilNr
'            End If
'            cFeld = Format$(dWert, "0")
'            cFeld = Space$(2 - Len(cFeld)) & cFeld
'            cLBSatz = cLBSatz & cFeld & " "
            
            List3.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    List3.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleListeArtikelWKL201"
    Fehler.gsFehlertext = "Im Programmteil Artikel drucken ist ein Fehler aufgetreten."

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


