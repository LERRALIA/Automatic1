VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL181 
   Caption         =   "automatischer Stammdatenabgleich"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL181.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
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
      Index           =   7
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   23
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Index           =   5
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   17
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   15
      Top             =   3840
      Width           =   615
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   12
      Top             =   3720
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
      Caption         =   "aktualisieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   11
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   5400
      MaxLength       =   6
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Index           =   0
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   6
      Left            =   9600
      TabIndex        =   3
      Top             =   2040
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
      Caption         =   "aktualisieren"
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
      Index           =   2
      Left            =   9600
      TabIndex        =   18
      Top             =   5160
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
      Caption         =   "aktualisieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   9600
      TabIndex        =   25
      Top             =   6720
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
      Caption         =   "aktualisieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   240
      TabIndex        =   26
      Top             =   7440
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Ihre LiefNr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   600
      TabIndex        =   24
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"frmWKL181.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   240
      TabIndex        =   22
      Top             =   6000
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   21
      Top             =   5280
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   $"frmWKL181.frx":04DD
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   240
      TabIndex        =   20
      Top             =   4440
      Width           =   7575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Farbnr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7320
      TabIndex        =   19
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Farbnr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7320
      TabIndex        =   16
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "aktualisiert bei allen Artikeln das Merkmal ""geräumt"" (Ex) und setzt ein Farbmerkmal bei allen betroffenen Artikeln."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   7575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Ihre LiefNr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Farbnr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7320
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Kisslive LiefNr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Ihre LiefNr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "aktualisiert alle Einkaufs- und Listenverkaufspreise eines Lieferantens und setzt ein Farbmerkmal bei allen veränderten Artikeln."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   2
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
      Caption         =   "automatischer Stammdatenabgleich"
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
      TabIndex        =   1
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "frmWKL181"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lcount As Long

    Select Case Index
        Case 0
            Unload frmWKL181
        Case 1
        
            lcount = Parf_EX_Aktualisieren(CLng(Text1(3).Text), Text1(4).Text)
            
            anzeige "normal", "Fertig! " & lcount & " Änderungen sind markiert worden.", Label1(4)
            
        Case 2
            
            lcount = Parf_Preise_Aktualisieren_ausExcel(Text1(5).Text)
            
            anzeige "normal", "Fertig! " & lcount & " Änderungen sind markiert worden.", Label1(4)
            
        Case 3 'Ex aus Excel
        
            lcount = EX_Aktualisieren_ausExcel(CLng(Text1(7).Text))
            
            anzeige "normal", "Fertig! " & lcount & " Änderungen sind markiert worden.", Label1(4)
        

        Case 6
            If Len(Text1(1).Text) = 6 Then
                lcount = Kisslive_Preise_Aktualisieren(CLng(Text1(0).Text), CLng(Text1(1).Text), Text1(2).Text)
            Else
                lcount = Parf_Preise_Aktualisieren(CLng(Text1(0).Text), Text1(2).Text)
            End If
            
            
            anzeige "normal", "Fertig! " & lcount & " Änderungen sind markiert worden.", Label1(4)
    End Select
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil AutoAbgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Parf_Preise_Aktualisieren_ausExcel(cAWM As String) As Long
On Error GoTo LOKAL_ERROR

    Parf_Preise_Aktualisieren_ausExcel = 0
    
    anzeige "normal", "Einstellungen werden überprüft...", Label1(4)
    
    LeseLIZENZ
    If gbLizenz = False Then
        Exit Function
    End If
    
    Dim sSQL As String
    Dim cPfad As String
    Dim dbExcel As Database
    Dim lAnzZ As Long
    Dim rsrs As Recordset
    Dim gsExcel50 As String
    Dim dKVK As Double
    Dim cKVK As String
    Dim bgefunden As Boolean
    
    bgefunden = True
    
    gsExcel50 = "Excel 5.0;"
    
    If pfadseekExcel_Artikel_im = False Then
        anzeige "rot2", "Abbruch durch Benutzer", Label1(4)
        Exit Function
    End If
    
    Screen.MousePointer = 11

    anzeige "normal", "", Label1(4)
    cPfad = Label1(10).Caption
    
    Set dbExcel = OpenDatabase(cPfad, 0, 0, gsExcel50)

    lAnzZ = 0
    Set rsrs = dbExcel.OpenRecordset("Abgleich$")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!EAN) Then
                
                dKVK = 0
                
                If Is_ArtNr_In(rsrs!EAN) Then
                
                    If Not IsNull(rsrs!UVP) Then
                        cKVK = rsrs!UVP
                    End If
                    
                    sSQL = "Update Artikel set KVKPR1 = " & cKVK
                    sSQL = sSQL & " , Lastdate = '" & DateValue(Now) & "' "
                    sSQL = sSQL & " , awm  = '" & cAWM & "' "
                    sSQL = sSQL & " where (ean = '" & rsrs!EAN & "' or ean2 = '" & rsrs!EAN & "' or ean3 = '" & rsrs!EAN & "')"
                    gdBase.Execute sSQL, dbFailOnError
                    
                    lAnzZ = lAnzZ + 1
                Else
                    bgefunden = False
                    schreibeProtokoll_Artikel_VKPREISE " " & rsrs!EAN & " nicht enthalten -> keine Aktualisierung"
                End If
                
                
                    
            End If
        rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If bgefunden = False Then
        zeigeHilfeDabapfad "LPROTOK", "Artikel_VKPREISE.txt"
    End If
        
    Screen.MousePointer = 0

    dbExcel.Close
    
    Parf_Preise_Aktualisieren_ausExcel = lAnzZ

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Parf_Preise_Aktualisieren_ausExcel"
    Fehler.gsFehlertext = "Im Programmteil AutoAbgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Function EX_Aktualisieren_ausExcel(lLinr As Long) As Long
On Error GoTo LOKAL_ERROR

    EX_Aktualisieren_ausExcel = 0
    
    anzeige "normal", "Einstellungen werden überprüft...", Label1(4)
    
    LeseLIZENZ
    If gbLizenz = False Then
        Exit Function
    End If
    
    Dim sSQL As String
    Dim cPfad As String
    Dim dbExcel As Database
    Dim lAnzZ As Long
    Dim rsrs As Recordset
    Dim gsExcel50 As String
    Dim sBESTELLNR As String
    Dim bgefunden As Boolean
    
    bgefunden = True
    
    gsExcel50 = "Excel 5.0;"
    
    If pfadseekExcel_Artikel_EX = False Then
        anzeige "rot2", "Abbruch durch Benutzer", Label1(4)
        Exit Function
    End If
    
    Screen.MousePointer = 11

    anzeige "normal", "", Label1(4)
    cPfad = Label1(14).Caption
    
    Set dbExcel = OpenDatabase(cPfad, 0, 0, gsExcel50)

    lAnzZ = 0
    Set rsrs = dbExcel.OpenRecordset("EXARTIKEL$")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BESTELLNR) Then
                
                sBESTELLNR = ""

                If Is_ArtNr_InARTLIEF(rsrs!BESTELLNR, lLinr) Then

                    If Not IsNull(rsrs!BESTELLNR) Then
                        sBESTELLNR = rsrs!BESTELLNR
                    End If

                    sSQL = "Update Artlief set RKZ = 'J' "
                    sSQL = sSQL & ", EXDAT = '" & DateValue(Now) & "' "
                    sSQL = sSQL & " where LINR = " & lLinr & " "
                    sSQL = sSQL & " and Libesnr = '" & sBESTELLNR & "' "
                    gdBase.Execute sSQL, dbFailOnError

                    lAnzZ = lAnzZ + 1
                Else
                    bgefunden = False
                    schreibeProtokoll_Artikel_EX " " & rsrs!BESTELLNR & " nicht enthalten -> keine Aktualisierung"
                End If
                
                
                    
            End If
        rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If bgefunden = False Then
        zeigeHilfeDabapfad "LPROTOK", "Artikel_EX.txt"
    End If
        
    Screen.MousePointer = 0

    dbExcel.Close
    
    EX_Aktualisieren_ausExcel = lAnzZ

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EX_Aktualisieren_ausExcel"
    Fehler.gsFehlertext = "Im Programmteil AutoAbgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Public Function Is_ArtNr_In(cEAN As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Is_ArtNr_In = False
    
    cSQL = "Select * from Artikel where (EAN = '" & cEAN & "' or EAN2 = '" & cEAN & "' or EAN3 = '" & cEAN & "')"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        Is_ArtNr_In = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Is_ArtNr_In"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function Is_ArtNr_InARTLIEF(cLiBesNr As String, lLinr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Is_ArtNr_InARTLIEF = False
    
    cSQL = "Select * from ARTLIEF where Libesnr = '" & cLiBesNr & "' and Linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        Is_ArtNr_InARTLIEF = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Is_ArtNr_InARTLIEF"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function pfadseekExcel_Artikel_im() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sExcelpfad  As String
    
    pfadseekExcel_Artikel_im = False

    sTitle = "Wo befindet sich die Datei?"
    
    
    sFilter = "Excel - Dateien (*.xls)|*.xls"
    
    sOldpfad = gcDBPfad & "\IN"
    sExcelpfad = pfadaendernplusDatname(sTitle, sFilter, sOldpfad)
    
    If sExcelpfad <> "" Then
        pfadseekExcel_Artikel_im = True
        Label1(10).Caption = sExcelpfad
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekExcel_Artikel_im"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function pfadseekExcel_Artikel_EX() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sExcelpfad  As String
    
    pfadseekExcel_Artikel_EX = False

    sTitle = "Wo befindet sich die Datei?"
    
    
    sFilter = "Excel - Dateien (*.xls)|*.xls"
    
    sOldpfad = gcDBPfad & "\IN"
    sExcelpfad = pfadaendernplusDatname(sTitle, sFilter, sOldpfad)
    
    If sExcelpfad <> "" Then
        pfadseekExcel_Artikel_EX = True
        Label1(14).Caption = sExcelpfad
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekExcel_Artikel_EX"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function Parf_Preise_Aktualisieren(lLinr As Long, cAWM As String) As Long
On Error GoTo LOKAL_ERROR

    Parf_Preise_Aktualisieren = 0
    
    anzeige "normal", "Einstellungen werden überprüft...", Label1(4)
    
    If gbOptiStada = False Then
        Exit Function
    End If
    
    
    
    LeseLIZENZ
    If gbLizenz = False Then
        Exit Function
    End If
    

    If Val(gcFilNr) > 0 Then
        Exit Function
    End If
    
    If IfOnline = "Offline" Then
        'MsgBox "Es besteht keine Verbindung zum Internet." & vbCrLf & "Bitte stellen Sie eine Online-Verbindung her und versuchen Sie es erneut."
        Exit Function
    End If
    
    
    
    If fTestLoginError = 0 Then 'ist alles OK? Datenbank erreichbar?
        'MsgBox "YES"
    Else
        'MsgBox "NONO"
        Exit Function
    End If

    Dim stConnect As String
    stConnect = "ODBC;DRIVER=SQL Server;SERVER=80.86.85.121;DATABASE=stada;UID=eanlive;PWD=sigverif2005"
'    stConnect = "ODBC;DRIVER=SQL Server;SERVER=80.86.85.121;DATABASE=stada;UID=sa;PWD=sigverif1"
    
    
    Dim dbKLive As DAO.Database

    Set dbKLive = OpenDatabase("STADA", dbDriverNoPrompt, False, stConnect)
    
    Dim sSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim rsRs2           As DAO.Recordset
    Dim rsKISS          As DAO.Recordset
    Dim lArtnrSQL       As Long
    Dim siVKPR          As Single
    Dim siLEKPR         As Single
    Dim lcount          As Long
    
    siVKPR = 0#
    siLEKPR = 0#
    lArtnrSQL = 0
    
    lcount = 0
    
    loeschNEW "AENPREIS", gdBase
    sSQL = " Create Table AENPREIS ("
    sSQL = sSQL & " ARTNR Long "
    sSQL = sSQL & " ,VKPR single "
    sSQL = sSQL & " ,LEKPR single "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Aktualisierung wird gestartet...", Label1(4)
    
    '1.
    sSQL = "Select Artikel.Artnr, Artikel.VKPR, Artlief.LEKPR from Artikel inner join Artlief "
    sSQL = sSQL & " on Artikel.artnr = Artlief.artnr "
    
    sSQL = sSQL & " where Artlief.Linr  = " & lLinr
    Set rsrs = dbKLive.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            siVKPR = 0#
            siLEKPR = 0#
            lArtnrSQL = 0
            
            anzeige "normal", lcount & " Artikel werden noch überprüft.", Label1(4)
    
            If Not IsNull(rsrs!artnr) Then
                lArtnrSQL = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                siVKPR = rsrs!vkpr
            End If
            
            If Not IsNull(rsrs!lekpr) Then
                siLEKPR = rsrs!lekpr
            End If
            
            sSQL = " Insert into AENPREIS ("
            sSQL = sSQL & " ARTNR "
            sSQL = sSQL & " ,VKPR "
            sSQL = sSQL & " ,LEKPR "
            sSQL = sSQL & " ) "
            sSQL = sSQL & " values "
            sSQL = sSQL & " ( "
            sSQL = sSQL & " " & lArtnrSQL & " "
            sSQL = sSQL & " ,'" & siVKPR & "' "
            sSQL = sSQL & " ,'" & siLEKPR & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
            lcount = lcount - 1
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "Preisabgleich wird gestartet...", Label1(4)
    
    sSQL = " Update Artlief inner join AENPREIS "
    sSQL = sSQL & " on artlief.artnr = aenpreis.artnr "
    sSQL = sSQL & " set artlief.lekpr = aenpreis.lekpr "
    sSQL = sSQL & " where artlief.linr = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update Artikel inner join AENPREIS "
    sSQL = sSQL & " on Artikel.artnr = aenpreis.artnr "
    sSQL = sSQL & " set Artikel.vkpr = aenpreis.vkpr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update Artikel inner join Artlief "
    sSQL = sSQL & " on Artikel.artnr = Artlief.artnr "
    sSQL = sSQL & " set Artikel.awm  = '" & cAWM & "' "
    sSQL = sSQL & " where artlief.linr = " & lLinr
    sSQL = sSQL & " and Round(Artikel.vkpr,2) <> Round(Artikel.KVKPR1,2) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    sSQL = "Select count(*) as maxi from Artikel inner join Artlief "
    sSQL = sSQL & " on Artikel.artnr = Artlief.artnr "
    sSQL = sSQL & " where artlief.linr = " & lLinr
    sSQL = sSQL & " and Artikel.awm  = '" & cAWM & "' "
    
    Set rsKISS = gdBase.OpenRecordset(sSQL)
    If Not rsKISS.EOF Then
    
        If Not IsNull(rsKISS!maxi) Then
            lcount = rsKISS!maxi
        End If

    End If
    rsKISS.Close

    Parf_Preise_Aktualisieren = lcount
    

    dbKLive.Close
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Parf_Preise_Aktualisieren"
    Fehler.gsFehlertext = "Im Programmteil AutoAbgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function Kisslive_Preise_Aktualisieren(lLinr As Long, lKissliveLinr As Long, cAWM As String) As Long
On Error GoTo LOKAL_ERROR

    Kisslive_Preise_Aktualisieren = 0
    
    anzeige "normal", "Einstellungen werden überprüft...", Label1(4)
    
    If gbOptiStadaSpiel = False Then
        Exit Function
    End If
    
    LeseLIZENZ_INDI
    If gbLizenzINDI = False Then
        Exit Function
    End If
        

    If Val(gcFilNr) > 0 Then
        Exit Function
    End If
    
    If IfOnline = "Offline" Then
        'MsgBox "Es besteht keine Verbindung zum Internet." & vbCrLf & "Bitte stellen Sie eine Online-Verbindung her und versuchen Sie es erneut."
        Exit Function
    End If
    
    
    
    If fTestLogin_Spiel_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
        'MsgBox "YES"
    Else
        'MsgBox "NONO"
        Exit Function
    End If

    Dim stConnect As String
    stConnect = "ODBC;DRIVER=SQL Server;SERVER=80.86.85.121;DATABASE=spielwaren;UID=eanlive;PWD=sigverif2005"
    
    
    Dim dbKLive As DAO.Database

    Set dbKLive = OpenDatabase("SPIELWAREN", dbDriverNoPrompt, False, stConnect)
    
    Dim sSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim rsRs2           As DAO.Recordset
    Dim rsKISS          As DAO.Recordset
    Dim lArtnrSQL       As Long
    Dim siVKPR          As Single
    Dim siLEKPR         As Single
    Dim cEAN            As String
    Dim lcount          As Long
    
    siVKPR = 0#
    siLEKPR = 0#
    lArtnrSQL = 0
    cEAN = ""
    
    lcount = 0
    
    loeschNEW "AENPREIS", gdBase
    
    sSQL = " Create Table AENPREIS ("
    sSQL = sSQL & " EAN Text(13) "
    sSQL = sSQL & " ,ARTNR long "
    sSQL = sSQL & " ,VKPR single "
    sSQL = sSQL & " ,LEKPR single "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Aktualisierung wird gestartet...", Label1(4)
    

    
'    sSQL = "select artean.ean, artean.artnr, 0 as vkpr, artlief.LEKPR "
'    sSQL = sSQL & " from artean inner join artlief on artean.artnr = artlief.artnr  "
'    sSQL = sSQL & " where artlief.linr = " & lKissliveLinr
    
    sSQL = "select artean.ean, artean.artnr,  artikel.vkpr, artlief.LEKPR "
    sSQL = sSQL & " from artean, artlief, artikel  "
    sSQL = sSQL & " where artean.artnr = artlief.artnr "
    sSQL = sSQL & " and artean.artnr = artikel.artnr "
    sSQL = sSQL & " and artlief.linr = " & lKissliveLinr
    
    Set rsrs = dbKLive.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            siVKPR = 0#
            siLEKPR = 0#
            lArtnrSQL = 0
            cEAN = ""
    
            If Not IsNull(rsrs!artnr) Then
                lArtnrSQL = rsrs!artnr
            End If
            
            anzeige "normal", lcount & " Artikel werden noch überprüft.", Label1(4)
            lcount = lcount - 1
            
'            'mit dieser Information holen wir den ListenVK
'            If lArtnrSQL > 0 Then
'                sSQL = "select vkpr from artikel where artnr = " & lArtnrSQL
'                Set rsRs2 = dbKLive.OpenRecordset(sSQL)
'                If Not rsRs2.EOF Then
'                    If Not IsNull(rsRs2!vkpr) Then
'                        siVKPR = rsRs2!vkpr
'                    End If
'
'                End If
'                rsRs2.Close
'            End If
            
            
            If Not IsNull(rsrs!EAN) Then
                cEAN = rsrs!EAN
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                siVKPR = rsrs!vkpr
            End If
            
            If Not IsNull(rsrs!lekpr) Then
                siLEKPR = rsrs!lekpr
            End If
            
            sSQL = " Insert into AENPREIS ("
            sSQL = sSQL & " ARTNR "
            sSQL = sSQL & " ,EAN "
            sSQL = sSQL & " ,VKPR "
            sSQL = sSQL & " ,LEKPR "
            sSQL = sSQL & " ) "
            sSQL = sSQL & " values "
            sSQL = sSQL & " ( "
            sSQL = sSQL & " " & 0 & " "
            sSQL = sSQL & " ,'" & cEAN & "' "
            sSQL = sSQL & " ,'" & siVKPR & "' "
            sSQL = sSQL & " ,'" & siLEKPR & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
'            lcount = lcount + 1
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "Preisabgleich wird gestartet...", Label1(4)
    
    sSQL = " Update AENPREIS inner join ARTIKEL "
    sSQL = sSQL & " on AENPREIS.EAN = ARTIKEL.EAN "
    sSQL = sSQL & " set AENPREIS.artnr = ARTIKEL.artnr "
    sSQL = sSQL & " where AENPREIS.artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update AENPREIS inner join ARTIKEL "
    sSQL = sSQL & " on AENPREIS.EAN = ARTIKEL.EAN2 "
    sSQL = sSQL & " set AENPREIS.artnr = ARTIKEL.artnr "
    sSQL = sSQL & " where AENPREIS.artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update AENPREIS inner join ARTIKEL "
    sSQL = sSQL & " on AENPREIS.EAN = ARTIKEL.EAN3 "
    sSQL = sSQL & " set AENPREIS.artnr = ARTIKEL.artnr "
    sSQL = sSQL & " where AENPREIS.artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    'Hier geht es los
    
    sSQL = " Update Artlief inner join AENPREIS "
    sSQL = sSQL & " on artlief.ARTNR = aenpreis.artnr "
    sSQL = sSQL & " set artlief.lekpr = aenpreis.lekpr "
    sSQL = sSQL & " where artlief.linr = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update Artikel inner join AENPREIS "
    sSQL = sSQL & " on Artikel.artnr = aenpreis.artnr "
    sSQL = sSQL & " set Artikel.vkpr = aenpreis.vkpr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update Artikel inner join Artlief "
    sSQL = sSQL & " on Artikel.artnr = Artlief.artnr "
    sSQL = sSQL & " set Artikel.awm  = '" & cAWM & "' "
    sSQL = sSQL & " where artlief.linr = " & lLinr
    sSQL = sSQL & " and Round(Artikel.vkpr,2) <> Round(Artikel.KVKPR1,2) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    sSQL = "Select count(*) as maxi from Artikel inner join Artlief "
    sSQL = sSQL & " on Artikel.artnr = Artlief.artnr "
    sSQL = sSQL & " where artlief.linr = " & lLinr
    sSQL = sSQL & " and Artikel.awm  = '" & cAWM & "' "
    
    Set rsKISS = gdBase.OpenRecordset(sSQL)
    If Not rsKISS.EOF Then
    
        If Not IsNull(rsKISS!maxi) Then
            lcount = rsKISS!maxi
        End If



    End If
    rsKISS.Close

    Kisslive_Preise_Aktualisieren = lcount
    

    dbKLive.Close
    
Exit Function
LOKAL_ERROR:

    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Kisslive_Preise_Aktualisieren"
        Fehler.gsFehlertext = "Im Programmteil AutoAbgleich ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    
    
End Function
Private Function Parf_EX_Aktualisieren(lLinr As Long, cAWM As String) As Long
On Error GoTo LOKAL_ERROR

    Parf_EX_Aktualisieren = 0
    
    anzeige "normal", "Einstellungen werden überprüft...", Label1(4)
    
    If gbOptiStada = False Then
        Exit Function
    End If
    
    
    
    LeseLIZENZ
    If gbLizenz = False Then
        Exit Function
    End If
    

    If Val(gcFilNr) > 0 Then
        Exit Function
    End If
    
    If IfOnline = "Offline" Then
        'MsgBox "Es besteht keine Verbindung zum Internet." & vbCrLf & "Bitte stellen Sie eine Online-Verbindung her und versuchen Sie es erneut."
        Exit Function
    End If
    
    
    
    If fTestLoginError = 0 Then 'ist alles OK? Datenbank erreichbar?
        'MsgBox "YES"
    Else
        'MsgBox "NONO"
        Exit Function
    End If

    Dim stConnect As String
    stConnect = "ODBC;DRIVER=SQL Server;SERVER=80.86.85.121;DATABASE=stada;UID=eanlive;PWD=sigverif2005"
    
    
    Dim dbKLive As DAO.Database

    Set dbKLive = OpenDatabase("STADA", dbDriverNoPrompt, False, stConnect)
    
    Dim sSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim rsRs2           As DAO.Recordset
    Dim rsKISS          As DAO.Recordset
    Dim lArtnrSQL       As Long
    Dim sEXDATUM        As String
    Dim sRKZ            As String
    Dim lcount          As Long
    
    sEXDATUM = "0"
    sRKZ = "N"
    lArtnrSQL = 0
    
    lcount = 0
    
    loeschNEW "AENEX", gdBase
    sSQL = " Create Table AENEX ("
    sSQL = sSQL & " ARTNR Long "
    sSQL = sSQL & " ,RKZ TEXT(1) "
    sSQL = sSQL & " ,EXDAT Datetime "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Aktualisierung wird gestartet...", Label1(4)
    
    '1.
    sSQL = "Select Artikel.Artnr, Artlief.RKZ, Artlief.EXDAT from Artikel inner join Artlief "
    sSQL = sSQL & " on Artikel.artnr = Artlief.artnr "
    
    sSQL = sSQL & " where Artlief.Linr  = " & lLinr
    sSQL = sSQL & " and Artlief.RKZ = true "
    Set rsrs = dbKLive.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            sRKZ = "N"
            sExdat = "0"
            
            lArtnrSQL = 0
            
            anzeige "normal", lcount & " Artikel werden noch überprüft.", Label1(4)
    
            If Not IsNull(rsrs!artnr) Then
                lArtnrSQL = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!RKZ) Then
                If rsrs!RKZ = True Then
                    sRKZ = "J"
                End If
            End If
            
            If Not IsNull(rsrs!EXDAT) Then
                sExdat = rsrs!EXDAT
            End If
            
            If sExdat = "" Then
                sExdat = "01.01.2010"
            End If
            
            If sExdat = "0" Then
                sExdat = "01.01.2010"
            End If
            
            sSQL = " Insert into AENEX ("
            sSQL = sSQL & " ARTNR "
            sSQL = sSQL & " ,RKZ "
            sSQL = sSQL & " ,EXDAT "
            sSQL = sSQL & " ) "
            sSQL = sSQL & " values "
            sSQL = sSQL & " ( "
            sSQL = sSQL & " " & lArtnrSQL & " "
            sSQL = sSQL & " ,'" & sRKZ & "' "
            
            
            sSQL = sSQL & " ," & CLng(DateValue(sExdat)) & " "
            
            
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
            lcount = lcount - 1
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "EXabgleich wird gestartet...", Label1(4)
    
    
    
    
    
    '2. Artikelfarbe setzen bei allen betroffenen Artikeln
    sSQL = "Select Artnr from AENEX "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            
            lArtnrSQL = 0
            
            anzeige "normal", lcount & " Artikel werden noch überprüft.", Label1(4)
    
            If Not IsNull(rsrs!artnr) Then
                lArtnrSQL = rsrs!artnr
            End If
            
            sSQL = " Update ARTIKEL inner join ARTLIEF "
            sSQL = sSQL & " on ARTIKEL.artnr = ARTLIEF.artnr "
            sSQL = sSQL & "  set Artikel.awm  = '" & cAWM & "' "
            sSQL = sSQL & " where ARTLIEF.RKZ = 'N' "
            sSQL = sSQL & " and ARTLIEF.LINR = " & lLinr & " "
            sSQL = sSQL & " and ARTIKEL.artnr = " & lArtnrSQL & " "
            gdBase.Execute sSQL, dbFailOnError
            
            
            
            
            
            
            
            lcount = lcount - 1
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "EXabgleich wird gestartet...", Label1(4)
    
'''''    sSQL = " Update ARTIKEL inner join AENEX "
'''''    sSQL = sSQL & " on ARTIKEL.artnr = AENEX.artnr "
''''''    sSQL = sSQL & " set ARTIKEL.RKZ = AENEX.RKZ "
''''''    sSQL = sSQL & " , ARTIKEL.EXDAT = AENEX.EXDAT "
'''''    sSQL = sSQL & "  set Artikel.awm  = '" & cAWM & "' "
'''''    sSQL = sSQL & " where ARTIKEL.RKZ = 'N' "
'''''
'''''    gdBase.Execute sSQL, dbFailOnError
'''''
    
    sSQL = " Update Artlief inner join AENEX "
    sSQL = sSQL & " on Artlief.artnr = AENEX.artnr "
    sSQL = sSQL & " set Artlief.RKZ = AENEX.RKZ "
    sSQL = sSQL & " , Artlief.EXDAT = AENEX.EXDAT "
'    sSQL = sSQL & " , Artikel.awm  = '" & cAWM & "' "
    sSQL = sSQL & " where Artlief.RKZ = 'N' "
    
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    

    sSQL = "Select count(*) as maxi from Artikel inner join Artlief "
    sSQL = sSQL & " on Artikel.artnr = Artlief.artnr "
    sSQL = sSQL & " where artlief.linr = " & lLinr
    sSQL = sSQL & " and Artikel.awm  = '" & cAWM & "' "

    Set rsKISS = gdBase.OpenRecordset(sSQL)
    If Not rsKISS.EOF Then

        If Not IsNull(rsKISS!maxi) Then
            lcount = rsKISS!maxi
        End If

    End If
    rsKISS.Close

    Parf_EX_Aktualisieren = lcount
    

    dbKLive.Close
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Parf_EX_Aktualisieren"
    Fehler.gsFehlertext = "Im Programmteil AutoAbgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil AutoAbgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW "AENPREIS", gdBase
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


