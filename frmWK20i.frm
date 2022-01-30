VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWK20i 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   2685
   ClientLeft      =   2565
   ClientTop       =   1005
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   2685
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin Threed.SSCommand SSCommand1 
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   " Ignorieren/Schließen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Bonus auszahlen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kundenangaben:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   10
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   8
      Left            =   2880
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   4
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   7
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   1
      X1              =   5760
      X2              =   5760
      Y1              =   120
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   0
      X1              =   2760
      X2              =   2760
      Y1              =   120
      Y2              =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusgrenze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmWK20i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cFirma          As String
Dim cVname          As String
Dim cNName          As String
Dim cTitel          As String
Dim cPlz            As String
Dim cStadt          As String
Dim cStrasse        As String
Dim FixBonus        As Double
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
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim derBonus        As Double
    
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
'    Modul6.alternativFarbform Me, Label1(0)
    
'    derBonus = ErmittleKundenBonus(gckundnr)
    
    cFirma = lookingForKundendaten(Trim(gckundnr)).firma
    cVname = lookingForKundendaten(Trim(gckundnr)).vorname
    cNName = lookingForKundendaten(Trim(gckundnr)).nachname
    cTitel = lookingForKundendaten(Trim(gckundnr)).titel
    cPlz = lookingForKundendaten(Trim(gckundnr)).Plz
    cStadt = lookingForKundendaten(Trim(gckundnr)).Ort
    cStrasse = lookingForKundendaten(Trim(gckundnr)).strasse
    
    Label1(2).Caption = "Bonusgrenze: " & Format(gdBonusGrenze, "#####0.00 EUR")
    Label1(3).Caption = "bonusfähiger Umsatz: " & Format(dBonusfaehig, "#####0.00 EUR")
    
    If gdBonusGrenze = 0 Then
        Label1(0).Caption = "0 x Bonus erreicht!"
        Label1(0).Refresh
    Else
        FixBonus = dBonusfaehig / gdBonusGrenze
        FixBonus = Fix(FixBonus)
        Label1(0).Caption = FixBonus & " x Bonus erreicht!"
        Label1(0).Refresh
    End If
    
    Label1(4).Caption = cFirma
    Label1(5).Caption = cTitel
    Label1(6).Caption = cVname
    Label1(7).Caption = cNName
    Label1(8).Caption = cPlz
    Label1(9).Caption = cStadt
    Label1(10).Caption = cStrasse
    
    schreibeProtokollUNITXT "Kundnr: " & gckundnr & " Bonusfenster angezeigt, " & FixBonus & " x Bonus erreicht!, " & Label1(3).Caption & " ", "BONUSANZEIGE"

    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil BonusAA auf. "
    
    Fehlermeldung1
End Sub
Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim Bonusneu        As String
    Dim Bonusalt        As String
    Dim rsrs            As Recordset
    Dim sSQL            As String
    Dim cDatum          As String
    Dim czeit           As String
    ReDim cZeilen(0 To 9) As String
    Dim lRet            As Long
    
    
    cDatum = DateValue(Now)
    czeit = TimeValue(Now)
    
    Select Case Index
        Case Is = 0     'Bonus reduzieren
            
            schreibeProtokollUNITXT "Kundnr: " & gckundnr & " Bonus auszahlen gedrückt durch Bediener: " & gcBediener, "BONUSANZEIGE"
            
'            gbbonusHerab = True
'            gbbonusausjetzt = True
'            gdbonusHerabwert = gdBonusGrenze * FixBonus


'            lRet = MsgBox("Möchten Sie wirklich den Bonus reduzieren?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage?")
'            If lRet = vbNo Then
'                Exit Sub
'            End If
            sSQL = "Select * from Kunden where kundnr = " & gckundnr
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                rsrs.Edit
                rsrs!Status = "E"
                rsrs!SYNStatus = "E"
                
                If Not IsNull(rsrs!BONUS) Then
                    Bonusalt = Format(rsrs!BONUS, "##000.00 EUR")
                    rsrs!BONUS = rsrs!BONUS - (FixBonus * gdBonusGrenze)
                    
                Else
                    Bonusalt = "000.00 EUR"
                    rsrs!BONUS = (FixBonus * gdBonusGrenze) * (-1)

                End If
                
                If Not IsNull(rsrs!TBONUS) Then
                    rsrs!TBONUS = CDbl(rsrs!TBONUS) - (FixBonus * gdBonusGrenze)
                Else
                    rsrs!TBONUS = (FixBonus * gdBonusGrenze) * (-1)
                End If
                
                rsrs!LASTDATE = DateValue(Now)
                rsrs!LASTTIME = TimeValue(Now)
                rsrs.Update
                
                Bonusneu = Format(dBonusfaehig - (FixBonus * gdBonusGrenze), "##000.00 EUR")

            End If
            
            rsrs.Close: Set rsrs = Nothing
            
            giAnzBonus_Erreicht = FixBonus
            Unload frmWK20i
                        
        Case Is = 2     'Bonus ignorieren

            schreibeProtokollUNITXT "Kundnr: " & gckundnr & " Ignorieren gedrückt durch Bediener: " & gcBediener, "BONUSANZEIGE"
            
            gbbonusHerab = False
            gbbonusausjetzt = False
            gdbonusHerabwert = 0
            
            giAnzBonus_Erreicht = 0
            Unload frmWK20i
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil BonusAA auf. "
    
    Fehlermeldung1
End Sub
