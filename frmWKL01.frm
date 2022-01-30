VERSION 5.00
Begin VB.Form frmWKL01 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6900
   ClientLeft      =   2715
   ClientTop       =   600
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   6900
   ScaleWidth      =   6060
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Speichern"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   5520
         Width           =   2295
      End
      Begin VB.TextBox Text3 
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
         Left            =   4800
         MaxLength       =   5
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox Text3 
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
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox Text3 
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
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox Text3 
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
         Left            =   120
         MaxLength       =   5
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   5040
         Width           =   1095
      End
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Abbrechen"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   3600
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "OK"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   4800
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   120
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   2520
         Width           =   1095
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
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1080
         Width           =   3495
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
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1080
         Width           =   975
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
         Left            =   1440
         MaxLength       =   35
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         Caption         =   "Achten Sie unbedingt auf Groß- und Kleinschreibung!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sie erhalten umgehend von uns den Bestätigungs-Code, der die Anwendung freischaltet."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   4200
         Width           =   5775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefon: (0511) 95 59 10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   3720
         Width           =   5775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         Caption         =   "Geben Sie bitte alle Angaben (Name, PLZ, Ort und die vier Zahlencodes) an KISS Hannover durch:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   5775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bestätigung durch KISS Warenwirtschaft GmbH, Hannover"
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
         Left            =   120
         TabIndex        =   14
         Top             =   4800
         Width           =   5775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         Caption         =   "Registrierungs-Code Ihrer Programmversion:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PLZ / Ort:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Name:"
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
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      Caption         =   "Programm-Registrierung "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmWKL01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function fnBerechneConfirmWerte() As Integer

    Dim cFeld(0 To 3) As String
    Dim cErg(0 To 3) As String
    Dim dWert1 As Double
    Dim dWert2 As Double
    Dim lWert As Long
    Dim cTmp As String

    fnBerechneConfirmWerte = 0
    
    cFeld(0) = Text2(0).Text
    cFeld(1) = Text2(1).Text
    cFeld(2) = Text2(2).Text
    cFeld(3) = Text2(3).Text

    '********************************************
    '* Berechnung Rückgabewert 1.Feld           *
    '********************************************

    'Feld 1: Kundenwert mal Länge(Kundenname) geteilt durch Wert der 5.Stelle bzw. durch 7

    cTmp = Text1(0).Text
    cTmp = LTrim$(RTrim$(cTmp))

    dWert1 = Val(cFeld(0))
    dWert2 = Len(cTmp)

    dWert1 = dWert1 * dWert2
    cTmp = LTrim$(RTrim$(Str$(dWert1)))
    If Len(cTmp) > 5 Then
        cTmp = Mid(cTmp, 5, 1)
    Else
        cTmp = "7"
    End If

    dWert2 = Val(cTmp)
    If dWert2 = 0 Then
        dWert2 = 7
    End If

    dWert1 = dWert1 / dWert2
    dWert1 = Fix(dWert1)
    cTmp = LTrim$(RTrim$(Str$(dWert1)))

    If Len(cTmp) > 5 Then
        cTmp = Left(cTmp, 5)
    Else
        cTmp = String$(5 - Len(cTmp), "0") + cTmp
    End If

    cErg(0) = cTmp

    
    '********************************************
    '* Berechnung Rückgabewert 2.Feld           *
    '********************************************

    'Feld 2: Kundenwert plus Postleitzahl geteilt durch Modulo9 bzw. 3 oder 7

    dWert1 = 0
    dWert2 = 0
    cTmp = ""

    cTmp = Text1(1).Text
    cTmp = LTrim$(RTrim$(cTmp))

    dWert1 = Val(cFeld(1))
    dWert2 = Val(cTmp)

    dWert1 = dWert1 + dWert2

    dWert2 = dWert1 Mod 9

    If dWert2 = 0 Then
        dWert2 = 3
    End If

    If dWert2 = 1 Then
        dWert2 = 7
    End If

    dWert1 = dWert1 / dWert2

    dWert1 = Fix(dWert1)

    cTmp = LTrim$(RTrim$(Str$(dWert1)))

    If Len(cTmp) > 5 Then
        cTmp = Left(cTmp, 5)
    Else
        cTmp = String$(5 - Len(cTmp), "0") + cTmp
    End If

    cErg(1) = cTmp


    '********************************************
    '* Berechnung Rückgabewert 3.Feld           *
    '********************************************

    'Feld 3: Kundenwert1 plus 2x Kundenwert2 plus 3xKundenwert3 geteilt durch 6

    dWert1 = 0
    dWert2 = 0
    cTmp = ""
     
    dWert1 = Val(cFeld(0))
    dWert2 = Val(cFeld(1))
    dWert2 = dWert2 * 2

    dWert1 = dWert1 + dWert2

    dWert2 = Val(cFeld(2))
    dWert2 = dWert2 * 3

    dWert1 = dWert1 + dWert2
    
    dWert1 = Fix(dWert1 / 6)

    cTmp = LTrim$(RTrim$(Str$(dWert1)))

    If Len(cTmp) > 5 Then
        cTmp = Left(cTmp, 5)
    Else
        cTmp = String$(5 - Len(cTmp), "0") + cTmp
    End If

    cErg(2) = cTmp

    '********************************************
    '* Berechnung Rückgabewert 4.Feld           *
    '********************************************

    'Feld 3: Kundenwert1 mod 9, Kundenwert2 mod 9, Kundenwert3 mod 9
    '        Modulowerte aufaddieren

    dWert1 = 0
    dWert2 = 0
    cTmp = ""
    
    dWert1 = Val(cFeld(0))
    dWert1 = dWert1 Mod 9
    dWert2 = dWert1

    dWert1 = Val(cFeld(1))
    dWert1 = dWert1 Mod 9
    dWert2 = dWert2 + dWert1

    dWert1 = Val(cFeld(2))
    dWert1 = dWert1 Mod 9
    dWert2 = dWert2 + dWert1

    dWert1 = Val(cFeld(3))
    dWert1 = dWert1 Mod 9
    dWert2 = dWert2 + dWert1

    dWert1 = Val(cFeld(3))
    dWert1 = dWert1 * 9
    dWert1 = dWert1 / dWert2
    dWert1 = Fix(dWert1)

    cTmp = LTrim$(RTrim$(Str$(dWert1)))

    If Len(cTmp) > 5 Then
        cTmp = Left(cTmp, 5)
    Else
        cTmp = String$(5 - Len(cTmp), "0") + cTmp
    End If
    
    cErg(3) = cTmp

    '************************************
    '* Auswertung                       *
    '************************************

    For lWert = 0 To 3
        cTmp = Text3(lWert).Text
        If cTmp <> cErg(lWert) Then
            fnBerechneConfirmWerte = lWert + 1
            Exit Function
        End If
    Next lWert
    

End Function

Private Sub BerechneRegisterWerteWKL01()
    On Error GoTo LOKAL_ERROR
    
    Dim cFirma As String
    Dim cAdresse As String
    Dim cSerienNr1 As String
    Dim cSerienNr2 As String
    
    Dim lcount As Long
    Dim strRoot As String
    Dim Seriennummer As Long
    Dim lngDummy As Long
    Dim strDummy As String
    
    
    cFirma = Text1(0).Text
    cFirma = Trim$(cFirma)
    cFirma = cFirma & Space$(35 - Len(cFirma))
    
    KonvertAnsiAscii cFirma
    
    lWert = 0
    For lcount = 1 To Len(cFirma)
        cZeichen = Mid(cFirma, lcount, 1)
        lWert = lWert + (Asc(cZeichen) * lcount)
    Next lcount
    Text2(0).Text = Format$(lWert, "#####00000")
    
    
    cAdresse = Text1(1).Text
    cAdresse = Trim$(cAdresse)
    cAdresse = cAdresse & Space$(7 - Len(cAdresse))
    
    cAdresse = cAdresse & Text1(2).Text
    cAdresse = Trim$(cAdresse)
    cAdresse = cAdresse & Space$(37 - Len(cAdresse))
    
    KonvertAnsiAscii cAdresse
    
    lWert = 0
    For lcount = 1 To Len(cAdresse)
        cZeichen = Mid(cAdresse, lcount, 1)
        lWert = lWert + (Asc(cZeichen) * lcount)
    Next lcount
    Text2(1).Text = Format$(lWert, "#####00000")
    

    strRoot = Left(App.Path, 3)
    GetVolumeInformation strRoot, strDummy, lngDummy, Seriennummer, lngDummy, lngDummy, strDummy, lngDummy
    cSerienNr1 = CStr(Seriennummer)

    If Len(cSerienNr1) > 5 Then
        cSerienNr2 = Right(cSerienNr1, 5)
        cSerienNr1 = Left(cSerienNr1, 5)
    Else
        cSerienNr2 = "04711"
    End If
    
    cSerienNr1 = String$(5 - Len(cSerienNr1), "0") & cSerienNr1
    cSerienNr2 = String$(5 - Len(cSerienNr2), "0") & cSerienNr2
    
    Text2(2).Text = cSerienNr1
    Text2(3).Text = cSerienNr2
    
Exit Sub
LOKAL_ERROR:
    MsgBox "frmWKL01.BerechneRegisterWerteWKL01: " & err.Number & " / " & err.Description
End Sub

Private Function fnPruefeDialogEingabeOKWKL01() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cTmp As String
    
    fnPruefeDialogEingabeOKWKL01 = 0
    
    cTmp = Text1(0).Text
    cTmp = Trim$(cTmp)
    If cTmp = "" Then
        fnPruefeDialogEingabeOKWKL01 = 1
        Exit Function
    End If
    
    cTmp = Text1(1).Text
    cTmp = Trim$(cTmp)
    If cTmp = "" Then
        fnPruefeDialogEingabeOKWKL01 = 2
        Exit Function
    End If
    
    cTmp = Text1(2).Text
    cTmp = Trim$(cTmp)
    If cTmp = "" Then
        fnPruefeDialogEingabeOKWKL01 = 3
        Exit Function
    End If
    
Exit Function
LOKAL_ERROR:
    MsgBox "frmfnPruefeDialogEingabeOKWKL01: " & err.Number & " / " & err.Description
End Function
Private Sub SchreibeRegistrierDateiWKL01()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim cPfad As String
    Dim cTmp As String
    Dim lcount As Long
    Dim cFeld As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cdatum As String
    Dim iFehler As Integer
    
    iFehler = 1
    
    cPfad = gcSysPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    iFehler = 2
    
    Kill cPfad & gcRegDatei
    
    cTmp = "MZ"
    
    For lcount = 1 To 68
        Randomize
        iWert = Int((254 * Rnd) + 128)
        If iWert > 255 Then
            iWert = iWert - 255
        End If
        cTmp = cTmp & Chr$(iWert)
    Next lcount
    
    iFehler = 3
    
    cTmp = cTmp & "This program cannot be run in DOS mode. $"
    
    For lcount = 1 To 100
        Randomize
        iWert = Int((254 * Rnd) + 128)
        If iWert > 255 Then
            iWert = iWert - 255
        End If
        cTmp = cTmp & Chr$(iWert)
    Next lcount
    
    iFehler = 4
    
    cFeld = ""
    cFeld = cFeld & Text1(0).Text & Chr$(27)
    cFeld = cFeld & Text1(1).Text & Chr$(27)
    cFeld = cFeld & Text1(2).Text & Chr$(27)
    cFeld = cFeld & Text2(0).Text & Chr$(27)
    cFeld = cFeld & Text2(1).Text & Chr$(27)
    cFeld = cFeld & Text2(2).Text & Chr$(27)
    cFeld = cFeld & Text2(3).Text & Chr$(27)
    cFeld = cFeld & Text3(0).Text & Chr$(27)
    cFeld = cFeld & Text3(1).Text & Chr$(27)
    cFeld = cFeld & Text3(2).Text & Chr$(27)
    cFeld = cFeld & Text3(3).Text & Chr$(27)
    cFeld = cFeld & Format$(Now, "DD.MM.YYYY") & Chr$(27)
    
    'MsgBox cFeld
    
    iFehler = 5
    
    cFeld = fnEncrypt(cFeld)
    
    iFehler = 6
    
    cTmp = cTmp & cFeld
    
    
    For lcount = 1 To 1958
        Randomize
        iWert = Int((254 * Rnd) + 128)
        If iWert > 255 Then
            iWert = iWert - 255
        End If
        cTmp = cTmp & Chr$(iWert)
    Next lcount

    iFehler = 7

    cdatum = Date
    Date = "01.01.2000"
    iFileNr = FreeFile
    If gbDebug Then
        MsgBox Trim$(Str$(iFileNr)) & " " & cPfad & gcRegDatei
    End If
    
'    MsgBox cPfad & gcRegDatei
    
    Open cPfad & gcRegDatei For Binary As #iFileNr
    Put #iFileNr, 1, cTmp
    Close iFileNr
    Date = cdatum
    
    iFehler = 8

    cSQL = "Select * from FIRMA"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    iFehler = 9

    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!name = Text1(0).Text
    rsrs!Plz = Text1(1).Text
    rsrs!Ort = Text1(2).Text
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    iFehler = 10

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        MsgBox "frmWKL01.SchreibeRegistrierDateiWKL01: " & err.Number & " / " & err.Description & " / Fehlerstufe = " & Trim$(Str$(iFehler))
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0
            iRet = fnPruefeDialogEingabeOKWKL01()
            Select Case iRet
                Case Is = 0
                    BerechneRegisterWerteWKL01
                    Text3(0).SetFocus
                Case Is = 1
                    MsgBox "Bitte geben Sie den Namen Ihrers Unternehmens an!", vbCritical, "STOP!"
                    Text1(0).SetFocus
                Case Is = 2
                    MsgBox "Bitte geben Sie die Postleitzahl an!", vbCritical, "STOP!"
                    Text1(1).SetFocus
                Case Is = 3
                    MsgBox "Bitte geben Sie den Ortsnamen an!", vbCritical, "STOP!"
                    Text1(2).SetFocus
            End Select
            
        Case Is = 1     'Beenden
            Unload frmWKL01
            
        Case Is = 2     'Speichern
            iRet = fnBerechneConfirmWerte()
            If iRet = 0 Then
                SchreibeRegistrierDateiWKL01
                Unload frmWKL01
            Else
                MsgBox "Die Bestätigungs-Codes stimmen nicht mit Ihren Registrierungsdaten überein!", vbCritical, "STOP!"
                Text3(iRet - 1).SetFocus
            End If
            
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    MsgBox "frmWKL01.Command1_Click: " & err.Number & " / " & err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    frmWKL01.Top = Screen.Height / 2 - frmWKL01.Height / 2
    frmWKL01.Left = Screen.Width / 2 - frmWKL01.Width / 2
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    Text2(3).Text = ""
    
    Text3(0).Text = ""
    Text3(1).Text = ""
    Text3(2).Text = ""
    Text3(3).Text = ""
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    MsgBox "frmWKL01.Form_Load: " & err.Number & " / " & err.Description
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyAscii <> 0 And KeyAscii <> 8 Then
        If Len(Text1(Index).Text) = Text1(Index).MaxLength - 1 Then
            If Index < 2 Then
                Text1(Index + 1).SetFocus
            Else
                Command1(0).SetFocus
            End If
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    MsgBox "frmWKL01.Text1_KeyPress: " & err.Number & " / " & err.Description
End Sub


Private Sub Text1_LostFocus(Index As Integer)
        Text1(Index).BackColor = vbWhite
End Sub


Private Sub Text3_GotFocus(Index As Integer)
    Text3(Index).BackColor = glSelBack1
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = Len(Text3(Index).Text)

End Sub


Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
    If KeyAscii <> 0 And KeyAscii <> 8 Then
        If Len(Text3(Index).Text) = Text3(Index).MaxLength - 1 Then
            If Index < 3 Then
                Text3(Index + 1).SetFocus
            Else
                Command1(2).SetFocus
            End If
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    MsgBox "frmWKL01.Text1_KeyPress: " & err.Number & " / " & err.Description

End Sub


Private Sub Text3_LostFocus(Index As Integer)
        Text3(Index).BackColor = vbWhite
End Sub


