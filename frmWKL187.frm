VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL187 
   BackColor       =   &H00C0C000&
   Caption         =   "Rabatt - Aufkleber"
   ClientHeight    =   8610
   ClientLeft      =   1215
   ClientTop       =   1590
   ClientWidth     =   11910
   Icon            =   "frmWKL187.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame3 
      Caption         =   "Sonderpreis - Aufkleber"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   11655
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "runde Etiketten ( DIN A4 100 Blatt á 40 Etiketten)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   5160
         TabIndex        =   27
         Top             =   1320
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "eckige Etiketten  (35,6 x 16,9 = 80 Stück/Blatt)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   5160
         TabIndex        =   26
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   240
         MaxLength       =   6
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   3
         Left            =   9480
         TabIndex        =   15
         Top             =   1080
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
      Begin VB.Label Label9 
         Caption         =   "Anzahl Etiketten"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Anzahl Leeretiketten"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Drucken Sie sich Sonderpreisetiketten, die nur den Preis auf dem Etikett enthalten."
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   8655
      End
      Begin VB.Label Label6 
         Caption         =   "Sonderpreis in €"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gesamtrabatt - Aufkleber"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   11655
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   240
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   2
         Left            =   9480
         TabIndex        =   10
         Top             =   1080
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
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "bestellen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   7560
         MouseIcon       =   "frmWKL187.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   25
         ToolTipText     =   "per Email bestellen"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Rote runde Klebeetiketten (DIN A4) 100 Blatt á 40 Etiketten für 32,75 Euro bestellbar bei KISS Hannover 0511 95510 (ArtNr 501721)"
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   23
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Label Label5 
         Caption         =   "Rabatt in %"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   $"frmWKL187.frx":074C
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   10215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Artikelrabatt - Aufkleber"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   11655
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   240
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   6
         Top             =   1080
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
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "bestellen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   7560
         MouseIcon       =   "frmWKL187.frx":0829
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   24
         ToolTipText     =   "per Email bestellen"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Rote runde Klebeetiketten (DIN A4) 100 Blatt á 40 Etiketten für 32,75 Euro bestellbar bei KISS Hannover 0511 95510 (ArtNr 501721)"
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   22
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   $"frmWKL187.frx":0B33
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   10215
      End
      Begin VB.Label Label2 
         Caption         =   "Rabatt in %"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7680
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Rabatt - Aufkleber"
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
      Width           =   11535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   840
      Y2              =   840
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
      TabIndex        =   1
      Top             =   7800
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL187"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
 On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case 0
            Unload frmWKL187
        Case 1
            If Val(Text1(0).Text) > 0 Then
                anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
                newRabattStrichcods Val(Text1(0).Text)
            Else
                anzeige "rot", "Geben Sie bitte die Rabatthöhe an!", Label1(4)
                Text1(0).SetFocus
            End If
        Case 2
            If Val(Text1(1).Text) > 0 Then
                anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
                newGesamtRabattStrichcods Val(Text1(1).Text)
            Else
                anzeige "rot", "Geben Sie bitte die Rabatthöhe an!", Label1(4)
                Text1(1).SetFocus
            End If
            
        Case 3
            If IsNumeric(Text1(2).Text) = True Then
                anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
                
                
                newSonderpreisEtiketten CDbl(Text1(2).Text), Val(Text1(3).Text), Val(Text1(4).Text), Option1(1).Value
            Else
                anzeige "rot", "Geben Sie bitte den Sonderpreis an!", Label1(4)
                Text1(2).SetFocus
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Rabatt - Aufkleber ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Rabatt - Aufkleber ist ein Fehler aufgetreten."
    
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

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    
    Label1(7).ForeColor = glS1
   
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    
    Label1(8).ForeColor = glS1
   
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 7 Then
        Etikettenbestellung_Per_Mail 501721, "Klebeetiketten DIN A4, rot rund, 1 Verpackungseinheit = 100 Blatt á 40 Etiketten = 4000 Etiketten"
    End If
    
    If Index = 8 Then
        Etikettenbestellung_Per_Mail 501721, "Klebeetiketten DIN A4, rot rund, 1 Verpackungseinheit = 100 Blatt á 40 Etiketten = 4000 Etiketten"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If Index = 7 Then
        Label1(7).ForeColor = glLink
    End If
    
    If Index = 8 Then
        Label1(8).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Rabatt - Aufkleber ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    
    Select Case Index
        Case 0
            cValid = "1234567890" & Chr$(8)
        Case 1, 3, 4
            cValid = "1234567890" & Chr$(8)
        Case 2
            cValid = "1234567890," & Chr$(8)
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
    Fehler.gsFehlertext = "Im Programmteil Rabatt - Aufkleber ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Rabatt - Aufkleber ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
