VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmWKL03 
   BackColor       =   &H00C0C000&
   Caption         =   "Etiketten aus Lieferscheinen"
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
   Icon            =   "frmWKL03.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   240
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   5880
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
   Begin sevCommand3.Command Command3 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   6480
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
   Begin sevCommand3.Command Command2 
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      Caption         =   "Tabelle leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   7920
      Width           =   3495
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      Caption         =   "zum Etikettendruck"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      Top             =   7920
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   5775
      Left            =   2880
      TabIndex        =   0
      Top             =   1440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10186
      _Version        =   393216
      FocusRect       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   7440
      Width           =   8775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferschein - Nr.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiketten aus Lieferscheinen"
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
      Width           =   11655
   End
End
Attribute VB_Name = "frmWKL03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aBreite() As Integer
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    If EtiErmittlung Then
        gsETILS = "aus Lieferschein"
    End If
    Unload frmWKL03
    frmWKL30.Show 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function EtiErmittlung() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL         As String
    Dim rsEtiLs      As Recordset
    
    EtiErmittlung = False
    
    sSQL = "Select * from LSTEETI "
    Set rsEtiLs = gdBase.OpenRecordset(sSQL)
    If Not rsEtiLs.EOF Then
        EtiErmittlung = True
    End If
    rsEtiLs.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EtiErmittlung"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR


    Dim sSQL        As String
    Dim bFound      As Boolean
    Dim lcount      As Long
    Dim cLiefschein As String

    bFound = False

    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount

    If Not bFound Then
        anzeige "rot", "Bitte markieren Sie einen Lieferschein!", lblanzeige
        Screen.MousePointer = 0
        Exit Sub
    End If

    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            cLiefschein = List1.list(lcount)

            sSQL = "Delete from ETIDRULS where LS = '" & cLiefschein & "' "
            gdBase.Execute sSQL, dbFailOnError

        End If
    Next lcount

    etilsAnzeigen

    LeseLieferschein_inListe "ETIDRULS"
    gsETILS = ""

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Command4_Click()
On Error GoTo LOKAL_ERROR
    
    MSHFLEX1.Clear

    etilsAnzeigen
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click()
    On Error GoTo LOKAL_ERROR
    
    
    Unload frmWKL03
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR
    
'    MSHFLEX1.Clear
    frmrefr
    FormatMShFlex1WKL03
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    frmrefr

    FormatMShFlex1WKL03
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
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
Private Sub LeseLieferschein_inListe(sTab As String)
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim cFeld As String

    List1.Clear
    'Tabellesuchen
    If sTab = "ETIDRULS" Then
        If NewTableSuchenDBKombi("ETIDRULS", gdBase) = False Then
            'Tabelle erstellen
            CreateTableT2 "ETIDRULS", gdBase
        End If
    End If

    cSQL = "Select distinct LS from  " & sTab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!LS) Then
                cFeld = rsrs!LS
                List1.AddItem cFeld
            Else
                cFeld = ""
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseLieferschein_inListe"
    Fehler.gsFehlertext = "Beim Ermitteln der Lieferscheinnummern ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub frmrefr()
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    LeseLieferschein_inListe "ETIDRULS"
    gsETILS = ""
    
    loeschNEW "LSTEETI", gdBase
    CreateTableT2 "LSTEETI", gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frmrefr"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMShFlex1WKL()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim i           As Integer
    Dim sSQL        As String
    
    sSQL = "Select * from LSTEETI order by LS "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    With MSHFLEX1
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            
            .Rows = lrow + 1
            .Row = lrow
            
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = 0
            End If
            
            .Col = 0
            .Text = lWert
            
            If Not IsNull(rsrs!BEZEICH) Then
                sWert = rsrs!BEZEICH
            Else
                sWert = ""
            End If
            
            .Col = 1
            .Text = sWert
            
            If Not IsNull(rsrs!vkpr) Then
                sWert = Format(rsrs!vkpr, "###,##0.00")
            Else
                sWert = 0
            End If
            
            .Col = 2
            .Text = sWert
            
            If Not IsNull(rsrs!ANZAHL) Then
                siWert = rsrs!ANZAHL
            Else
                siWert = 0
            End If
            
            .Col = 3
            .Text = siWert
            
            If Not IsNull(rsrs!filnr) Then
                siWert = rsrs!filnr
            Else
                siWert = 0
            End If
            
            .Col = 4
            .Text = siWert
            
            If Not IsNull(rsrs!LS) Then
                sWert = rsrs!LS
            Else
                sWert = ""
            End If
            
            .Col = 5
            .Text = sWert
            
            For i = 0 To 4
                If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
                    aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                End If
            Next i
            
            rsrs.MoveNext
        Loop
    End If
    
    For i = 0 To 4
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.2
    Next i
    
    rsrs.Close: Set rsrs = Nothing
    
    .RowHeight(1) = 0
    lrow = lrow - 1

    .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex1WKL03"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub etilsAnzeigen()
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsEtiLs     As Recordset
    Dim bFound      As Boolean
    Dim lcount      As Long
    Dim cLiefschein As String

    loeschNEW "LSTEETI", gdBase
    CreateTableT2 "LSTEETI", gdBase

    bFound = False

    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount

    If Not bFound Then
        anzeige "rot", "Bitte markieren Sie einen Lieferschein!", lblanzeige
        Screen.MousePointer = 0
        Exit Sub
    End If

    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            cLiefschein = List1.list(lcount)

            sSQL = "Delete from LSTEETI where LS = '" & cLiefschein & "' "
            gdBase.Execute sSQL, dbFailOnError

            sSQL = "Insert into LSTEETI Select * from etidruls where LS = '" & cLiefschein & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
    Next lcount

    FormatMShFlex1WKL03
    FuellenMShFlex1WKL

    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak

    Command2.Visible = True

 Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "etilsAnzeigen"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSHFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
    Dim i           As Integer
    Dim j           As Integer
    
    With gridx
    
        ReDim bBreit(.Cols - 1)
        
        For j = 0 To .Rows - 1
            For i = 0 To .Cols - 1
                If TextWidth(.TextMatrix(j, i)) > bBreit(i) Then
                    bBreit(i) = TextWidth(.TextMatrix(j, i))
                End If
            Next i
        Next j
        
        Select Case Screen.Height
            Case Is > 15000
                siFak = 1.5
            Case Is > 12000
                siFak = 1.4
            Case Is > 11000
                siFak = 1.2
            Case Is > 10000
                siFak = 1.1
            Case Is > 8000
                siFak = 1#
        End Select
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = bBreit(i) * siFak * siEigFak
        Next i
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FormatMShFlex1WKL03()
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer
    
    With MSHFLEX1
        .Clear
        
        .Rows = 25
        .Cols = 6
         ReDim aBreite(.Cols)
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .Text = "Artnr"
        
        .Col = 1
        .Text = "Artikelbezeichnung"
        
        .Col = 2
        .Text = "KVK - Preis"
        
        .Col = 3
        .Text = "Anzahl/Etiketten"
        
        .Col = 4
        .Text = "Filiale"
        
        .Col = 5
        .Text = "Lieferschein"
        
        For j = 0 To .Cols - 1
            .Col = j
            aBreite(j) = TextWidth(.TextMatrix(0, j))
        Next j
        
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatMShFlex1WKL03"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_DblClick()
    On Error GoTo LOKAL_ERROR

    sortierenHGrid MSHFLEX1
      
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Etiketten aus Lieferscheinen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
