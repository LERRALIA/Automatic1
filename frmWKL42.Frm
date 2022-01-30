VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL42 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Lieferantenliste"
   ClientHeight    =   8625
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   11910
   Icon            =   "frmWKL42.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   480
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   8415
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   4
         Left            =   6000
         TabIndex        =   18
         Top             =   6720
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   3
         Left            =   1920
         TabIndex        =   10
         Top             =   6720
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
         Caption         =   "Stückzahlen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   6720
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
         Caption         =   "Umsatz"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   1
         Left            =   10200
         TabIndex        =   11
         Top             =   6720
         Width           =   1575
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   0
         Left            =   4200
         TabIndex        =   8
         Top             =   6720
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
         Caption         =   "Drucken"
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
         Height          =   6150
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   11655
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   11655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   8280
         TabIndex        =   15
         Top             =   6960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   8880
         TabIndex        =   14
         Top             =   6720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "von"
         Height          =   255
         Index           =   1
         Left            =   8400
         TabIndex        =   13
         Top             =   6720
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   12
         Top             =   6720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
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
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11655
      Begin VB.CheckBox Check1 
         Caption         =   "nur aktive Lieferanten"
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   480
         Width           =   2535
      End
      Begin sevCommand3.Command Command1 
         Height          =   435
         Index           =   1
         Left            =   9360
         TabIndex        =   4
         Top             =   1200
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   435
         Index           =   0
         Left            =   9360
         TabIndex        =   3
         Top             =   240
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
         Caption         =   "Suche"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "nach Lieferanten-Nummer"
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "nach Alphabet"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   4455
      End
      Begin sevCommand3.Command Command1 
         Height          =   435
         Index           =   12
         Left            =   9360
         TabIndex        =   20
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   767
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
            Size            =   9.75
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
         Caption         =   "Vergleich"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lblUeberschrift 
         BackStyle       =   0  'Transparent
         Caption         =   "Sortierung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferantenliste"
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
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   8055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
        
    Screen.MousePointer = 11
        
    Select Case Index
        Case Is = 0
            LeseDatenWKL42
            
            Frame2.Visible = True
            Frame1.Enabled = False
        Case Is = 1
            Unload frmWKL42
        Case Is = 12
            frmWKL214.Show 1
            Me.Refresh
        
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "ARTHEAD", gdBase
    loeschNEW "ARTDRUCK", gdBase
    
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
Private Sub LeseDatenWKL42()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    ReDim cOrderBy(0 To 1) As String
    Dim lcount As Long
    Dim lOrderBy As Long
    Dim ctmp As String
    Dim cEkPr As String
    Dim dWert As Double
    Dim cFeld As String
    Dim cLBSatz As String
    
    
    cOrderBy(0) = "order by LIEFBEZ "
    cOrderBy(1) = "order by LINR "
    
    
    For lcount = 0 To 1
        If Option1(lcount).Value = True Then
            lOrderBy = lcount
            Exit For
        End If
    Next lcount
    
    cSQL = "Delete from LISRT where linr is null "
    gdBase.Execute cSQL, dbFailOnError
    cSQL = "Delete from LISRT where linr = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "select * from lisrt where( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )"
    If Check1.Value = vbChecked Then
        cSQL = cSQL & " and not kuerzel is null and kuerzel <> '' "
    End If
    cSQL = cSQL & cOrderBy(lOrderBy)
    
    List1.Clear
    List2.Clear
    
    
    ctmp = "LiefNr Lieferanten-Bezeichnung           PLZ     Stadt                          Straße                         Telefon              Fax"
    List1.AddItem ctmp
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                cFeld = rsrs!linr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 6 Then
                MsgBox "Ungültige Feldlänge bei LINR", vbCritical, "STOP!"
                Exit Sub
            End If
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!LIEFBEZ) Then
                cFeld = rsrs!LIEFBEZ
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            If InStr(cFeld, Chr$(39)) > 0 Then
                Mid(cFeld, InStr(cFeld, Chr$(39)), 1) = "´"
            End If
            If Len(cFeld) > 30 Then
                MsgBox "Ungültige Feldlänge bei LIEFBEZ", vbCritical, "STOP!"
                Exit Sub
            End If
            cFeld = cFeld & Space$(30 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " " & Space$(3)
            
            If Not IsNull(rsrs!Plz) Then
                cFeld = rsrs!Plz
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 7 Then
                MsgBox "Ungültige Feldlänge bei PLZ", vbCritical, "STOP!"
                Exit Sub
            End If
            cFeld = cFeld & Space$(7 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!STADT) Then
                cFeld = rsrs!STADT
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 30 Then
                MsgBox "Ungültige Feldlänge bei STADT", vbCritical, "STOP!"
                Exit Sub
            End If
            cFeld = cFeld & Space$(30 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!strasse) Then
                cFeld = rsrs!strasse
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 30 Then
                MsgBox "Ungültige Feldlänge bei STADT", vbCritical, "STOP!"
                Exit Sub
            End If
            cFeld = cFeld & Space$(30 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!Tel) Then
                cFeld = rsrs!Tel
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            If Len(cFeld) < 20 Then
                cFeld = cFeld & Space$(20 - Len(cFeld))
            Else
                cFeld = Left(cFeld, 20)
            End If
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!Fax) Then
                cFeld = rsrs!Fax
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            If Len(cFeld) < 20 Then
                cFeld = cFeld & Space$(20 - Len(cFeld))
            Else
                cFeld = Left(cFeld, 20)
            End If
            cLBSatz = cLBSatz & cFeld & " "
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        Label1(2).Caption = List2.ListCount
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseDatenWKL42"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Select Case Index
        Case Is = 0     'Ergebnisliste drucken
            DruckeErgebnisListeWKL42
        Case Is = 1     'Ergebnisliste schließen
            Frame1.Enabled = True
            Frame2.Visible = False
        Case 2 To 3     'Umsätze oder Stückzahlen
            If List2.ListIndex < 0 Then
                MsgBox "Bitte einen Lieferanten markieren!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
                Exit Sub
            End If
            Label1(3).Caption = Trim$(Str$(Index))
            frmWK40c.Show 1
        Case 4     'Excelexport
        
            loeschNEW "LIEFEXPORT", gdBase
            CreateTableT2 "LIEFEXPORT", gdBase
            
            cSQL = "Insert into LIEFEXPORT Select "
            cSQL = cSQL & " LINR  "
            cSQL = cSQL & ", LIEFBEZ "
            cSQL = cSQL & ", KUERZEL "
            cSQL = cSQL & ", PLZ "
            cSQL = cSQL & ", STADT "
            cSQL = cSQL & ", STRASSE "
            cSQL = cSQL & ", TEL "
            cSQL = cSQL & ", FAX "
            cSQL = cSQL & ", ZUSATZ "
            cSQL = cSQL & ", KUNDNR  "
            cSQL = cSQL & ", NOTIZ  "
            cSQL = cSQL & " from LISRT "
            gdBase.Execute cSQL, dbFailOnError
            
            ExcelExport "LIEFEXPORT", gdBase
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub DruckeErgebnisListeWKL42()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim lcount As Long
    Dim lAnzSatz As Long
    Dim cLBSatz As String
    
    loeschNEW "ARTHEAD", gdBase
    loeschNEW "ARTDRUCK", gdBase

    cSQL = "Create Table ARTHEAD (SCHLUESSEL DOUBLE, LISTTEXT TEXT(180))"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create INDEX SCHLUESSEL on ARTHEAD (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Table ARTDRUCK (SCHLUESSEL DOUBLE, LISTTEXT TEXT(180))"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create INDEX SCHLUESSEL on ARTDRUCK (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cLBSatz = List1.list(0)
    cSQL = "Insert into ARTHEAD (SCHLUESSEL, LISTTEXT) values (1, '" & cLBSatz & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    lAnzSatz = List2.ListCount
    
    Label1(2).Caption = Format$(lAnzSatz, "###,##0")
    Label1(2).Refresh
    For lcount = 0 To lAnzSatz - 1
        Label1(0).Caption = Format$(lcount + 1, "###,##0")
        Label1(0).Refresh
        cLBSatz = List2.list(lcount)
        cSQL = "Insert into ARTDRUCK (SCHLUESSEL, LISTTEXT) values (1, '" & cLBSatz & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next lcount
    
    reportbildschirm "WKL002", "aWKL42"

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 3372 Or err.Number = 3376 Or err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeErgebnisListeWKL42"
        Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten. "
        
        Fehlermeldung1

    End If
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift(0)
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten. "
    
    Fehlermeldung1

End Sub
Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    Frame2.Top = 840
    Frame2.Left = 0
    Frame2.Height = 7455
    Frame2.Width = 11895
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenliste ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub


