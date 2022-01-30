VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL164 
   Caption         =   "Willkommen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL164.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame6 
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
      Height          =   9945
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   11895
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   0
         Left            =   6720
         TabIndex        =   13
         Top             =   5040
         Width           =   3615
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
         Caption         =   "weiter ohne Service"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   120
         MaxLength       =   25
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   2400
         Width           =   4335
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   5
         Top             =   5040
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
         Caption         =   "Senden"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   120
         MaxLength       =   35
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   120
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   5040
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   120
         MaxLength       =   25
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   3240
         Width           =   4335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Sie können sich natürlich auch telefonisch für die 30 Tage Unterstützung anmelden. 0511/9559110"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   5760
         Width           =   11295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Ansprechpartner:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Telefonnummer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Registrieren Sie sich und Sie erhalten neben der freien Programmversion auch die ersten 30 Tage kostenlose Unterstützung."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   11415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Ihre Emailadresse:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "Willkommen bei K.I.S.S. Hannover"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   10215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Firma:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmWKL164"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click(index As Integer)
On Error GoTo LOKAL_ERROR

Select Case index
    Case 0
        Unload frmWKL164
    Case 2
'        sende Email

        Dim cAbsenderEmail As String
        Dim cAnEmailadresse As String
        Dim cBetreff As String
        Dim cMessagetext As String
        Dim sAttachment As String
        Dim sFirma As String
        Dim sBetreff As String
        Dim cAnsprech As String
        Dim cTelefon As String
        
        sAttachment = ""
        cTelefon = Text3(4).Text
        sFirma = Text3(5).Text
        cAnsprech = Text3(1).Text
        cAbsenderEmail = Text3(0).Text
        cAnEmailadresse = "vertrieb@kisswws.de"
        cBetreff = "Registrierung für WINKISS FREE"
        
        cMessagetext = "Sehr geehrtes KISS Team, " & vbCrLf & vbCrLf
        cMessagetext = cMessagetext & "hiermit möchte ich meine Winkiss free Version bei Ihnen registrieren." & vbCrLf & vbCrLf
        cMessagetext = cMessagetext & "Ansprechpartner: " & cAnsprech & vbCrLf
        cMessagetext = cMessagetext & "Firma: " & sFirma & vbCrLf
        cMessagetext = cMessagetext & "Email: " & cAbsenderEmail & vbCrLf
        cMessagetext = cMessagetext & "Telefon: " & cTelefon & vbCrLf
        
        cMessagetext = cMessagetext & "freundliche Grüße"
        
        schickeMailimHintergrundSSL sFirma, cAbsenderEmail, "", cAnEmailadresse _
    , "vertrieb@kisswws.de", gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, cBetreff, cMessagetext, sAttachment
    
    
        CreateTableT2 "FREE", gdBase
        


        Unload frmWKL164
End Select
            
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel fehlt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    WKL145Positionieren
    Modul6.Skalieren Me, True, True:
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    Screen.MousePointer = 0
    
    Text3(0).Text = ""
    Text3(1).Text = ""
    Text3(4).Text = ""
    Text3(5).Text = ""

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikel fehlt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKL145Positionieren()
On Error GoTo LOKAL_ERROR
    
    Frame6.Top = 0
    Frame6.Left = 0
    Frame6.Height = 9000
    Frame6.Width = 12000
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL145Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Artikel fehlt ist ein Fehler aufgetreten."
    
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
Private Sub Text3_Change(index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case index
    
        Case 0
        
            Command4(2).Enabled = False
            If Len(Text3(0).Text) > 4 Then
                If InStr(Text3(0).Text, "@") > 0 Then
                    If InStr(Text3(0).Text, ".") > 0 Then
                        Command4(2).Enabled = True
                    End If
                End If
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_Change"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text3_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(index).BackColor = glSelBack1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Text3_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command4_Click 2
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

