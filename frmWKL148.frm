VERSION 5.00
Begin VB.Form frmWKL148 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Lagerwerte"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame5"
      Height          =   6735
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
         Height          =   4860
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   11295
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
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   11295
      End
      Begin VB.Label Label4 
         Caption         =   "Pennerwert zum Schnitteinkaufspreis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   13
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "LW SEK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "PW SEK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Lagerwert zum Schnitteinkaufspreis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   10
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label Label4 
         Caption         =   "Pennerbestand in Stück"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Lagerbestand in Stück"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   8
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Pbestand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Lbestand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
   End
   Begin sevCommand3.Command Command3 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Zurück"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
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
      Width           =   9135
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Lagerwerte"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   0
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
Attribute VB_Name = "frmWKL148"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    With Frame5
        .Height = 6735
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
    Fehler.gsFehlertext = "Im Programmteil Lagerwerte ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL148
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Lagerwerte ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren

    Schrift Me
    alternativFarbform Me, lblUeberschrift
    LogtoStart Me
    
    List1.AddItem "Monat                     Lbestand   LW SEK     Pbestand   PW SEK     PAnteil am SEK in %"
    If Left(gcSuch, 4) = "LINR" Then
        LeseLagerwerte
        
    ElseIf Left(gcSuch, 4) = "LPZX" Then
        LeseLagerwerte
        
    ElseIf Left(gcSuch, 4) = "MARK" Then
        LeseLagerwerte
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Lagerwerte ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseLagerwerte()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim cFeld   As String
    Dim cLBSatz As String
    Dim rsrs    As Recordset
    Dim rsrs1    As Recordset
    Dim clinr   As String
    Dim clpz    As String
    Dim cMarkenbez As String
    Dim dLSEK   As Double
    Dim dPSEK   As Double
    Dim dAnteil   As Double
    
    If Left(gcSuch, 4) = "LINR" Then
        clinr = Right(gcSuch, Len(gcSuch) - 4)
        loeschNEW "LP" & srechnertab, gdBase
        loeschNEW "LP1" & srechnertab, gdBase
        
        cSQL = "Select "
        cSQL = cSQL & " YEAR(DATUM) as JAHR"
        cSQL = cSQL & ", MONTH(DATUM) as MONAT"
        cSQL = cSQL & ", LINR"
        cSQL = cSQL & ", AVG(SEK) as AVGSEK"
        cSQL = cSQL & ", AVG(BEST) as AVGBEST"
        cSQL = cSQL & ", 0.0 as AVGPSEK"
        cSQL = cSQL & ", 0.0 as AVGPBEST"
        cSQL = cSQL & " into LP" & srechnertab & " from LAGERLW"
        cSQL = cSQL & " where Linr = " & clinr
        cSQL = cSQL & " group by  YEAR(DATUM), MONTH(DATUM), LINR "
        cSQL = cSQL & " order by  YEAR(DATUM) desc, MONTH(DATUM) desc "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select "
        cSQL = cSQL & " YEAR(DATUM) as JAHR"
        cSQL = cSQL & ", MONTH(DATUM) as MONAT"
        cSQL = cSQL & ", LINR"
        cSQL = cSQL & ", AVG(SEK) as AVGPSEK"
        cSQL = cSQL & ", AVG(BEST) as AVGPBEST"
        cSQL = cSQL & " into LP1" & srechnertab & " from PENLAGERLW"
        cSQL = cSQL & " where Linr = " & clinr
        cSQL = cSQL & " group by  YEAR(DATUM), MONTH(DATUM), LINR "
        cSQL = cSQL & " order by  YEAR(DATUM) desc, MONTH(DATUM) desc "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update LP" & srechnertab & " inner join LP1" & srechnertab
        cSQL = cSQL & " on LP" & srechnertab & ".LINR = LP1" & srechnertab & ".LINR "
        cSQL = cSQL & " and LP" & srechnertab & ".JAHR = LP1" & srechnertab & ".JAHR "
        cSQL = cSQL & " and LP" & srechnertab & ".MONAT = LP1" & srechnertab & ".MONAT "
        cSQL = cSQL & " set LP" & srechnertab & ".AVGPSEK = LP1" & srechnertab & ".AVGPSEK "
        cSQL = cSQL & " , LP" & srechnertab & ".AVGPBEST = LP1" & srechnertab & ".AVGPBEST "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select "
        cSQL = cSQL & "  JAHR"
        cSQL = cSQL & ", MONAT"
        cSQL = cSQL & ", AVGSEK"
        cSQL = cSQL & ", AVGBEST"
        cSQL = cSQL & ", AVGPSEK"
        cSQL = cSQL & ", AVGPBEST"
        cSQL = cSQL & " from LP" & srechnertab
        cSQL = cSQL & " order by Jahr desc, Monat desc "
        anzeige "normal", ermLiefBez(CLng(clinr)), lblanzeige
        
    ElseIf Left(gcSuch, 4) = "LPZX" Then
        clpz = Right(gcSuch, Len(gcSuch) - 4)
        
        loeschNEW "LP" & srechnertab, gdBase
        loeschNEW "LP1" & srechnertab, gdBase
        
        cSQL = "Select "
        cSQL = cSQL & " YEAR(DATUM) as JAHR"
        cSQL = cSQL & ", MONTH(DATUM) as MONAT"
        cSQL = cSQL & ", LINR"
        cSQL = cSQL & ", LPZ"
        cSQL = cSQL & ", AVG(SEK) as AVGSEK"
        cSQL = cSQL & ", AVG(BEST) as AVGBEST"
        cSQL = cSQL & ", 0.0 as AVGPSEK"
        cSQL = cSQL & ", 0.0 as AVGPBEST"
        cSQL = cSQL & " into LP" & srechnertab & " from LAGERLLW"
        cSQL = cSQL & " where Linr = " & gclinr
        cSQL = cSQL & " and LPZ = " & clpz
        cSQL = cSQL & " group by  YEAR(DATUM), MONTH(DATUM), LINR ,lpz"
        cSQL = cSQL & " order by  YEAR(DATUM) desc, MONTH(DATUM) desc "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select "
        cSQL = cSQL & " YEAR(DATUM) as JAHR"
        cSQL = cSQL & ", MONTH(DATUM) as MONAT"
        cSQL = cSQL & ", LINR"
        cSQL = cSQL & ", LPZ"
        cSQL = cSQL & ", AVG(SEK) as AVGPSEK"
        cSQL = cSQL & ", AVG(BEST) as AVGPBEST"
        cSQL = cSQL & " into LP1" & srechnertab & " from PENLAGERLLW"
        cSQL = cSQL & " where Linr = " & gclinr
        cSQL = cSQL & " and LPZ = " & clpz
        cSQL = cSQL & " group by  YEAR(DATUM), MONTH(DATUM), LINR,lpz "
        cSQL = cSQL & " order by  YEAR(DATUM) desc, MONTH(DATUM) desc "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update LP" & srechnertab & " inner join LP1" & srechnertab
        cSQL = cSQL & " on LP" & srechnertab & ".LINR = LP1" & srechnertab & ".LINR "
        cSQL = cSQL & " and LP" & srechnertab & ".LPZ = LP1" & srechnertab & ".LPZ "
        cSQL = cSQL & " and LP" & srechnertab & ".JAHR = LP1" & srechnertab & ".JAHR "
        cSQL = cSQL & " and LP" & srechnertab & ".MONAT = LP1" & srechnertab & ".MONAT "
        cSQL = cSQL & " set LP" & srechnertab & ".AVGPSEK = LP1" & srechnertab & ".AVGPSEK "
        cSQL = cSQL & " , LP" & srechnertab & ".AVGPBEST = LP1" & srechnertab & ".AVGPBEST "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select "
        cSQL = cSQL & "  JAHR"
        cSQL = cSQL & ", MONAT"
        cSQL = cSQL & ", AVGSEK"
        cSQL = cSQL & ", AVGBEST"
        cSQL = cSQL & ", AVGPSEK"
        cSQL = cSQL & ", AVGPBEST"
        cSQL = cSQL & " from LP" & srechnertab
        cSQL = cSQL & " order by Jahr desc, Monat desc "
        
        anzeige "normal", ermLINBEZ1(CLng(clpz), CLng(gclinr)), lblanzeige
    ElseIf Left(gcSuch, 4) = "MARK" Then
    
        loeschNEW "LP" & srechnertab, gdBase
        loeschNEW "LP1" & srechnertab, gdBase
        cMarkenbez = Right(gcSuch, Len(gcSuch) - 4)
        
        cSQL = "Select "
        cSQL = cSQL & " YEAR(DATUM) as JAHR"
        cSQL = cSQL & ", MONTH(DATUM) as MONAT"
        cSQL = cSQL & ", AVG(SEK) as AVGSEK"
        cSQL = cSQL & ", AVG(BEST) as AVGBEST"
        cSQL = cSQL & ", 0.0 as AVGPSEK"
        cSQL = cSQL & ", 0.0 as AVGPBEST"
        cSQL = cSQL & ", '" & cMarkenbez & "' as Marke "
        cSQL = cSQL & " into LP" & srechnertab & " from LAGERMW "
        cSQL = cSQL & " where Marke = '" & cMarkenbez & "'"
        cSQL = cSQL & " group by  YEAR(DATUM), MONTH(DATUM),Marke "
        cSQL = cSQL & " order by  YEAR(DATUM) desc, MONTH(DATUM) desc "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select "
        cSQL = cSQL & " YEAR(DATUM) as JAHR"
        cSQL = cSQL & ", MONTH(DATUM) as MONAT"
        cSQL = cSQL & ", AVG(SEK) as AVGPSEK"
        cSQL = cSQL & ", AVG(BEST) as AVGPBEST"
        cSQL = cSQL & ", '" & cMarkenbez & "' as Marke "
        cSQL = cSQL & " into LP1" & srechnertab & " from PENLAGERMW "
        cSQL = cSQL & " where Marke = '" & cMarkenbez & "'"
        cSQL = cSQL & " group by  YEAR(DATUM), MONTH(DATUM),Marke "
        cSQL = cSQL & " order by  YEAR(DATUM) desc, MONTH(DATUM) desc "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update LP" & srechnertab & " inner join LP1" & srechnertab
        cSQL = cSQL & " on LP" & srechnertab & ".MARKE = LP1" & srechnertab & ".MARKE "
        cSQL = cSQL & " and LP" & srechnertab & ".JAHR = LP1" & srechnertab & ".JAHR "
        cSQL = cSQL & " and LP" & srechnertab & ".MONAT = LP1" & srechnertab & ".MONAT "
        cSQL = cSQL & " set LP" & srechnertab & ".AVGPSEK = LP1" & srechnertab & ".AVGPSEK "
        cSQL = cSQL & " , LP" & srechnertab & ".AVGPBEST = LP1" & srechnertab & ".AVGPBEST "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select "
        cSQL = cSQL & "  JAHR"
        cSQL = cSQL & ", MONAT"
        cSQL = cSQL & ", AVGSEK"
        cSQL = cSQL & ", AVGBEST"
        cSQL = cSQL & ", AVGPSEK"
        cSQL = cSQL & ", AVGPBEST"
        cSQL = cSQL & " from LP" & srechnertab
        cSQL = cSQL & " order by Jahr desc, Monat desc "

        anzeige "normal", cMarkenbez, lblanzeige
    End If
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cLBSatz = ""
            If Not IsNull(rsrs!Monat) Then
                cFeld = gcMonat(rsrs!Monat)
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(10 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "

            If Not IsNull(rsrs!jahr) Then
                cFeld = rsrs!jahr
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(5 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & Space(10)
            
            If Not IsNull(rsrs!AVGBEST) Then
                cFeld = rsrs!AVGBEST
            Else
                cFeld = ""
            End If
            cFeld = Format$(cFeld, "######0")
            cFeld = cFeld & Space$(10 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "

            If Not IsNull(rsrs!AVGSEK) Then
                cFeld = rsrs!AVGSEK
                dLSEK = cFeld
            Else
                cFeld = ""
                dLSEK = 0
            End If
            cFeld = Format$(cFeld, "########0.00")
            cFeld = cFeld & Space$(10 - Len(cFeld))
            
            
            cLBSatz = cLBSatz & cFeld & " "
            
            
            
            
            
            
            If Not IsNull(rsrs!AVGPBEST) Then
                cFeld = rsrs!AVGPBEST
            Else
                cFeld = ""
            End If
            cFeld = Format$(cFeld, "######0")
            cFeld = cFeld & Space$(10 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "

            If Not IsNull(rsrs!AVGPSEK) Then
                cFeld = rsrs!AVGPSEK
                dPSEK = cFeld
            Else
                cFeld = ""
                dPSEK = 0
            End If
            cFeld = Format$(cFeld, "########0.00")
            cFeld = cFeld & Space$(10 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If dLSEK <> 0 Then
                dAnteil = dPSEK * 100 / dLSEK
            Else
                dAnteil = 0
            End If
            
            cFeld = Format$(dAnteil, "########0.00")
            cFeld = cFeld & Space$(12 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            List3.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseLagerwerte"
    Fehler.gsFehlertext = "Im Programmteil Lagerwerte ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "BE" & srechnertab, gdBase
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



