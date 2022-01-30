VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL122 
   Caption         =   "Sonderkontingente"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL122.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   7440
      TabIndex        =   25
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
      Height          =   3840
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   11535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   11535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   20
      Top             =   2520
      Width           =   1095
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   5280
      TabIndex        =   12
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
      Caption         =   "Entfernen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   11
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   2
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
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Meldemenge:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   0
      TabIndex        =   22
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "wenn diese Menge unterschitten wird, dann Alarm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   2640
      TabIndex        =   21
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Artikel:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   9600
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Listenverkaufspreis:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   7320
      TabIndex        =   17
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   9600
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Kassenpreis zur Zeit:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7320
      TabIndex        =   15
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "hier den Sonderpreis vormerken"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   14
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "hier Sonderkontingentstückzahl eingeben"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   13
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9600
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Bestand zur Zeit:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Preis:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Bezeich"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Stück:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Artnr:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
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
      TabIndex        =   4
      Top             =   7920
      Width           =   9255
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonderkontingente"
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
      Width           =   9135
   End
End
Attribute VB_Name = "frmWKL122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL122
        Case 1
            speicherdaten
        Case 2
            Deldaten
        Case 3
            DruckenKontin
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    anzeige "normal", "", Label1(4)
    
    zeigedaten gsARTNR
    FuelleListeKontin
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigedaten(sArt As String)
On Error GoTo LOKAL_ERROR
    
    Dim rsrs    As Recordset
    Dim sSQL    As String
    
    Label4(0).Caption = ""
    
    If sArt = "" Then
        Exit Sub
    End If
    
    If IsNumeric(sArt) = False Then
        Exit Sub
    End If
    
    sSQL = " select * from artikel where artnr = " & sArt
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                Label4(0).Caption = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                Label4(2).Caption = rsrs!BEZEICH
            End If
            
            
            If Not IsNull(rsrs!BESTAND) Then
                Label4(5).Caption = rsrs!BESTAND
            End If
            
            If Not IsNull(rsrs!KVKPR1) Then
                Label4(9).Caption = Format(rsrs!KVKPR1, "######0.00")
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                Label4(11).Caption = Format(rsrs!vkpr, "######0.00")
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
    Fehler.gsFunktion = "zeigedaten"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleListeKontin()
On Error GoTo LOKAL_ERROR
    
    Dim rsrs    As Recordset
    Dim sSQL    As String
    Dim cSatz   As String
    Dim cFeld   As String
    
    Dim lIstbestand     As Long
    Dim lABestand       As Long
    Dim lVerfuegbar     As Long
    Dim lSondermenge    As Long
    Dim lMeldemenge     As Long
    
    Dim lAnz            As Long
    
    lAnz = 0
    
    List1.Clear
    List2.Clear
    List2.Refresh
    
    List1.AddItem "Artnr   Bezeichnung                          Datum     Uhrzeit   Sondermenge   noch verfügbar  "
    
    '(artnr,BEZEICH,menge,lastdate,lasttime,WPREIS,MeldeMenge,ABESTAND,KVKPR1)
    sSQL = " select * from Kontin "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            lAnz = lAnz + 1
            cSatz = ""
            cFeld = ""
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            End If
            cSatz = cFeld
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH & Space(35 - Len(rsrs!BEZEICH))
            Else
                cFeld = Space(35)
            End If
            cSatz = cSatz & "  " & cFeld
            
            If Not IsNull(rsrs!LASTDATE) Then
                cFeld = Format(rsrs!LASTDATE, "DD.MM.YY")
            Else
                cFeld = Space(8)
            End If
            cSatz = cSatz & "  " & cFeld
            
            If Not IsNull(rsrs!LASTTIME) Then
                cFeld = Format(rsrs!LASTTIME, "hh:mm:ss")
            Else
                cFeld = Space(8)
            End If
            cSatz = cSatz & "  " & cFeld
            
            If Not IsNull(rsrs!Menge) Then
                cFeld = rsrs!Menge & Space(12 - Len(rsrs!Menge))
                lSondermenge = rsrs!Menge
            Else
                cFeld = Space(6)
                lSondermenge = 0
            End If
            cSatz = cSatz & "  " & cFeld
            
        
            lIstbestand = ermBESTAND(rsrs!artnr)
            
            If Not IsNull(rsrs!ABESTAND) Then
                lABestand = rsrs!ABESTAND
            Else
                lABestand = 0
            End If
            
            If Not IsNull(rsrs!Meldemenge) Then
                lMeldemenge = rsrs!Meldemenge
            Else
                lMeldemenge = 0
            End If

            lVerfuegbar = lIstbestand - lABestand '- lSondermenge
            
            cSatz = cSatz & "  " & lVerfuegbar & Space(18 - Len(CStr(lVerfuegbar)))
            
            If lMeldemenge > lVerfuegbar Then
                cSatz = cSatz & "Achtung! Meldemenge erreicht"
            End If
            
            List2.AddItem cSatz
                
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lAnz = 0 Then
        anzeige "normal", "", Label1(4)
    ElseIf lAnz = 1 Then
        anzeige "normal", lAnz & " Artikel wird angezeigt", Label1(4)
    Else
        anzeige "normal", lAnz & " Artikel werden angezeigt", Label1(4)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleListeKontin"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherdaten()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim sArtnr      As String
    Dim sPreis      As String
    Dim sKVKPR1     As String
    Dim lMenge      As Long
    Dim lBestandA   As Long
    Dim lMeldemenge As Long
    Dim iRet        As Integer
    Dim ctmp        As String
    
    If Label4(0).Caption = "" Then
        anzeige "normal", "", Label1(4)
        Exit Sub
    End If
    
    If Text1(2).Text = "" Then
         Text1(2).Text = "0"
    End If
    
    If Text1(1).Text = "" Then
         Text1(1).Text = "0,00"
    End If
    
    If Trim(Text1(1).Text) = "," Then
         Text1(1).Text = "0,00"
    End If
    
    If Text1(0).Text = "" Then
         Text1(0).Text = "0"
    End If
    
    lMeldemenge = Text1(2).Text
    lMenge = Text1(0).Text
    sPreis = Text1(1).Text
    lBestandA = Label4(5).Caption
    sKVKPR1 = Label4(9).Caption
    sArtnr = Label4(0).Caption
    
    If lBestandA > 0 Then
        If lBestandA >= lMenge Then
            ctmp = "Ist die Sondermenge von " & lMenge & vbCrLf & vbCrLf
            ctmp = ctmp & "schon im jetzigen Bestand von " & lBestandA & " enthalten?"
            iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbYes Then
               lBestandA = lBestandA - lMenge
            End If
        End If
    End If
    
    sSQL = "Insert into KONTIN (artnr,BEZEICH,menge,lastdate,lasttime,WPREIS,MeldeMenge,ABESTAND,KVKPR1) values  "
    sSQL = sSQL & " ( " & sArtnr & " ,'" & Label4(2).Caption & "', " & lMenge
    sSQL = sSQL & ", '" & DateValue(Now) & "'"
    sSQL = sSQL & ", '" & TimeValue(Now) & "'"
    sSQL = sSQL & ", '" & sPreis & "'"
    sSQL = sSQL & ", " & lMeldemenge
    sSQL = sSQL & ", " & lBestandA
    sSQL = sSQL & ", '" & sKVKPR1 & "'"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    Label4(0).Caption = ""
    anzeige "normal", "erfolgreich gespeichert", Label1(4)
    
    FuelleListeKontin
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherdaten"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Deldaten()
On Error GoTo LOKAL_ERROR

    Dim cLBSatz         As String
    Dim cArtNr          As String
    Dim cBestelltam     As String
    Dim cBestelltum     As String
    Dim cSQL            As String
    
    If List2.ListIndex < 0 Then
        MsgBox "Bitte einen Eintrag auswählen!", vbInformation, "Winkiss Hinweis:"
        List2.SetFocus
        Exit Sub
    End If
    
    cLBSatz = List2.list(List2.ListIndex)
    cBestelltam = Mid$(cLBSatz, 46, 8)
    cBestelltum = Mid$(cLBSatz, 56, 8)
    cArtNr = Left$(cLBSatz, 6)
        
    cSQL = "Delete from KONTIN   "
    cSQL = cSQL & " where ARTNR = " & cArtNr
    cSQL = cSQL & " and Lastdate = " & CLng(DateValue(cBestelltam))
    cSQL = cSQL & " and lasttime = '" & cBestelltum & "'"
    gdBase.Execute cSQL, dbFailOnError
    
   
    anzeige "normal", "", Label1(4)

    FuelleListeKontin
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Deldaten"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckenKontin()
On Error GoTo LOKAL_ERROR

    Dim cSQL            As String
    
    If Not Datendrin("KONTIN", gdBase) Then
        anzeige "rot", "keine Druckdaten vorhanden", Label1(4)
        Exit Sub
    End If
    
    loeschNEW "KONTINP", gdBase
    CreateTable "KONTINP", gdBase
    
    cSQL = "Insert into KONTINP Select * from KONTIN   "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update KONTINP inner join ARTIKEL on KONTINP.ARTNR = ARTIKEL.ARTNR "
    cSQL = cSQL & " set KONTINP.KVKPR1 = ARTIKEL.KVKPR1 "
    cSQL = cSQL & " , KONTINP.BESTAND = ARTIKEL.BESTAND"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update KONTINP "
    cSQL = cSQL & " set Verfueg = BESTAND - ABESTAND "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
    
    reportbildschirm "", "awkl122"
    
    anzeige "normal", "", Label1(4)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckenKontin"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    Select Case Index
        Case 0, 2
            cValid = "1234567890" & Chr(8)
            cZeichen = Chr$(KeyAscii)
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 1
            cValid = "1234567890," & Chr(8)
            cZeichen = Chr$(KeyAscii)
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command5_Click 1
    End If
    
    
    If KeyCode = vbKeyEscape Then
        Command5_Click 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Sonderkontingente ist ein Fehler aufgetreten."
    
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
