VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK25c 
   Caption         =   "Zeitenstatistik"
   ClientHeight    =   8595
   ClientLeft      =   1140
   ClientTop       =   1800
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
   Icon            =   "frmWK25c.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check1 
      Caption         =   "nur umsatzf‰hige"
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
      Left            =   3840
      TabIndex        =   24
      Top             =   3000
      Width           =   2055
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   9600
      TabIndex        =   21
      Top             =   840
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
      Caption         =   "Heute"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.CheckBox Check1 
      Caption         =   "kumulierte Daten"
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
      Left            =   3840
      TabIndex        =   20
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Height          =   1335
      Left            =   3720
      TabIndex        =   15
      Top             =   720
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "1 Stunde"
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
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "30 Minuten"
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
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Ansicht unterteilen"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sonntag"
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
      Index           =   1
      Left            =   6480
      TabIndex        =   14
      Top             =   3480
      Value           =   1  'Aktiviert
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Samstag"
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
      Left            =   6480
      TabIndex        =   13
      Top             =   3120
      Value           =   1  'Aktiviert
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Freitag"
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
      Index           =   6
      Left            =   6480
      TabIndex        =   12
      Top             =   2760
      Value           =   1  'Aktiviert
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Donnerstag"
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
      Left            =   6480
      TabIndex        =   11
      Top             =   2400
      Value           =   1  'Aktiviert
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mittwoch"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   2040
      Value           =   1  'Aktiviert
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dienstag"
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
      Left            =   6480
      TabIndex        =   9
      Top             =   1680
      Value           =   1  'Aktiviert
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Montag"
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
      Left            =   6480
      TabIndex        =   7
      Top             =   1320
      Value           =   1  'Aktiviert
      Width           =   2415
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   6
      Top             =   2040
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
      Caption         =   "Schlieﬂen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   3
      Top             =   1440
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
      Caption         =   "Suche Daten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   20
      Left            =   2640
      TabIndex        =   22
      ToolTipText     =   "Kalender"
      Top             =   720
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      ToolTip         =   "W‰hlen Sie hier das Datum aus."
      ToolTipTitle    =   "Kalender"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   21
      Left            =   2640
      TabIndex        =   23
      ToolTipText     =   "Kalender"
      Top             =   1200
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      ToolTip         =   "W‰hlen Sie hier das Datum aus."
      ToolTipTitle    =   "Kalender"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   7320
      Width           =   11415
   End
   Begin VB.Label Label2 
      Caption         =   "nur bestimmte Wochentage betrachten"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6480
      TabIndex        =   8
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Datum bis:"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Datum von:"
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
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Zeitenstatistik"
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
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmWK25c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPruef As Integer

Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 20        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 21        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim cMM As String
    Dim cYYYY As String
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    cMM = Month(Now)
    cYYYY = Year(Now)
    
    cMM = String$(2 - Len(cMM), "0") & cMM
    
    Text1(0).Text = "01." & cMM & "." & cYYYY
    Text1(1).Text = Format$(Now, "dd.mm.yyyy")
    
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
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
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim iFileNr As Integer
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case 0 'Suche daten
            
            If Check1(0).Value = vbUnchecked Then
            
                If erstelleDaten Then
                    reportbildschirm "", "aWKL25ca"
                End If
            Else
                If erstelleDaten Then
                    reportbildschirm "", "aWKL25cb"
                End If
            End If
             
        Case 2  'Schlieﬂen
            loeschNEW "zeitzone", gdBase
            Unload frmWK25c
        Case 3 'heute
        
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
            
            Option1(1).Value = True
            
            If erstelleDaten Then
                reportbildschirm "", "aWKL25ca"
            End If
        
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub intervall(b30 As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sIntervall(46)  As String
    Dim sSQL            As String
    
    sIntervall(0) = "00:00:00"
    sIntervall(1) = "00:30:00"
    sIntervall(2) = "01:00:00"
    sIntervall(3) = "01:30:00"
    sIntervall(4) = "02:00:00"
    sIntervall(5) = "02:30:00"
    sIntervall(6) = "03:00:00"
    sIntervall(7) = "03:30:00"
    sIntervall(8) = "04:00:00"
    sIntervall(9) = "05:30:00"
    sIntervall(10) = "06:00:00"
    sIntervall(11) = "06:30:00"
    sIntervall(12) = "07:00:00"
    sIntervall(13) = "07:30:00"
    sIntervall(14) = "08:00:00"
    sIntervall(15) = "08:30:00"
    sIntervall(16) = "09:00:00"
    sIntervall(17) = "09:30:00"
    sIntervall(18) = "10:00:00"
    sIntervall(19) = "10:30:00"
    sIntervall(20) = "11:00:00"
    sIntervall(21) = "11:30:00"
    sIntervall(22) = "12:00:00"
    sIntervall(23) = "12:30:00"
    sIntervall(24) = "13:00:00"
    sIntervall(25) = "13:30:00"
    sIntervall(26) = "14:00:00"
    sIntervall(27) = "14:30:00"
    sIntervall(28) = "15:00:00"
    sIntervall(29) = "15:30:00"
    sIntervall(30) = "16:00:00"
    sIntervall(31) = "16:30:00"
    sIntervall(32) = "17:00:00"
    sIntervall(33) = "17:30:00"
    sIntervall(34) = "18:00:00"
    sIntervall(35) = "18:30:00"
    sIntervall(36) = "19:00:00"
    sIntervall(37) = "19:30:00"
    sIntervall(38) = "20:00:00"
    sIntervall(39) = "20:30:00"
    sIntervall(40) = "21:00:00"
    sIntervall(41) = "21:30:00"
    sIntervall(42) = "22:00:00"
    sIntervall(43) = "22:30:00"
    sIntervall(44) = "23:00:00"
    sIntervall(45) = "23:30:00"
    
    loeschNEW "Zeitintv", gdBase
    
    sSQL = "create Table zeitintv ( "
    sSQL = sSQL & " Timeint Text(8) "
    sSQL = sSQL & " , preis single "
    sSQL = sSQL & " , kanz single "
    sSQL = sSQL & " , menge single  "
    
    sSQL = sSQL & " , nse single  "
    sSQL = sSQL & " , nsp single  "
    
    
    
    sSQL = sSQL & " , nspAbs single "
    sSQL = sSQL & " , nseAbs single "
    sSQL = sSQL & " , preisAbs single "
    sSQL = sSQL & " , kanzAbs single "
    sSQL = sSQL & " , mengeAbs single ) "
    gdBase.Execute sSQL, dbFailOnError
    
    If b30 Then
        For i = 0 To 45
            sSQL = "Insert into zeitintv (timeint) values ('" & sIntervall(i) & "')"
            gdBase.Execute sSQL, dbFailOnError
        Next i
    Else
        For i = 0 To 45 Step 2
            sSQL = "Insert into zeitintv (timeint) values ('" & sIntervall(i) & "')"
            gdBase.Execute sSQL, dbFailOnError
        Next i
    End If
        
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890." & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    If cZeichen = "," Then
        cZeichen = "."
        KeyAscii = Asc(cZeichen)
    End If
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = 13 Then
        Command1_Click 0
    End If
    If KeyCode = 27 Then
        Command1_Click 2
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function erstelleDaten() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim lVon            As Long
    Dim lBis            As Long
    Dim i               As Integer
    Dim j               As Integer
    Dim timevon         As String
    Dim timebis         As String
    Dim rsInt           As Recordset
    Dim rszeitz         As Recordset
    Dim rsKass          As Recordset
    Dim sAuswertzr      As String
    Dim sAuswertkr      As String
    Dim alleWeekdays    As Boolean
    Dim lanztage        As Long
    
    erstelleDaten = False
    
    Label3.ForeColor = vbRed
    Label3.Caption = datumspruefung(Text1(0).Text, Text1(1).Text)
    Label3.Refresh
        
    If Label3.Caption <> "" Then
        Exit Function
    End If
    
    Label3.ForeColor = glS1
    Label3.Caption = "Daten werden ermittelt..."
    Label3.Refresh

    Screen.MousePointer = 11
    
    Me.Refresh
    
    
    Screen.MousePointer = 11
    
    intervall Option1(0).Value
    
    lVon = DateValue(Text1(0).Text)
    lBis = DateValue(Text1(1).Text)
    sAuswertzr = Text1(0).Text & " - " & Text1(1).Text
    sAuswertkr = "alle Wochentage"
    alleWeekdays = True
    
    loeschNEW "Zeitzone", gdBase
    
    sSQL = "Create Table zeitzone ( "
    sSQL = sSQL & " aTime datetime "
    sSQL = sSQL & " , preis single "
    sSQL = sSQL & " , adate datetime "
    sSQL = sSQL & " , belegnr single "
    sSQL = sSQL & " , Kasnum BYTE "
    sSQL = sSQL & " , menge long "
    sSQL = sSQL & " , NSE single "
    sSQL = sSQL & " , NSP single "
    sSQL = sSQL & "  ) "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "zhead", gdBase
    
    sSQL = "create Table zhead ( "
    sSQL = sSQL & " auszr text(24) "
    sSQL = sSQL & " , auskr text(254) "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into zeitzone select azeit as aTime "
    sSQL = sSQL & " , preis, menge, adate, belegnr,kasnum "
    sSQL = sSQL & " , ((((Preis/(100 + " & gdMWStE & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStE & "))* 100) as NSP"
    sSQL = sSQL & " , (Preis * 100) /(100 + " & gdMWStE & ") - (EKPR * Menge) as NSE"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & "  where ADATE BETWEEN " & lVon & " "
    sSQL = sSQL & "  and " & lBis & " "
    If Check1(8).Value = vbChecked Then
        sSQL = sSQL & "  and UMS_OK = 'J' "
    End If
    sSQL = sSQL & "  and ARTNR <> 666666 "
    sSQL = sSQL & "  and Filiale = " & gcFilNr & " "
    sSQL = sSQL & " and MWST = 'E'"
    sSQL = sSQL & "  and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into zeitzone select azeit as aTime "
    sSQL = sSQL & " , preis, menge, adate, belegnr,kasnum "
    sSQL = sSQL & " ,((((A.Preis/(100 + " & gdMWStV & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStV & "))* 100) as NSP"
    sSQL = sSQL & " ,(A.Preis * 100) /(100 + " & gdMWStV & ") - (A.EKPR * A.Menge) as NSE"
    sSQL = sSQL & " from Kassjour A "
    sSQL = sSQL & "  where ADATE BETWEEN " & lVon & " "
    sSQL = sSQL & "  and " & lBis & " "
    If Check1(8).Value = vbChecked Then
        sSQL = sSQL & "  and UMS_OK = 'J' "
    End If
    sSQL = sSQL & "  and ARTNR <> 666666 "
    sSQL = sSQL & "  and Filiale = " & gcFilNr & " "
    sSQL = sSQL & " and MWST = 'V'"
    sSQL = sSQL & "  and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into zeitzone select azeit as aTime "
    sSQL = sSQL & " , preis, menge, adate, belegnr,kasnum "
    sSQL = sSQL & " ,((((A.Preis/(100 + " & gdMWStO & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStO & "))* 100) as NSP"
    sSQL = sSQL & " ,(A.Preis * 100) /(100 + " & gdMWStO & ") - (A.EKPR * A.Menge) as NSE"
    sSQL = sSQL & " from Kassjour A "
    sSQL = sSQL & "  where ADATE BETWEEN " & lVon & " "
    sSQL = sSQL & "  and " & lBis & " "
    
    If Check1(8).Value = vbChecked Then
        sSQL = sSQL & "  and UMS_OK = 'J' "
    End If
    
    sSQL = sSQL & "  and ARTNR <> 666666 "
    sSQL = sSQL & "  and Filiale = " & gcFilNr & " "
    sSQL = sSQL & " and MWST = 'O'"
    sSQL = sSQL & "  and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    For i = 1 To 7
        If Check1(i).Value = vbUnchecked Then
            sSQL = "Delete from zeitzone "
            sSQL = sSQL & " where weekday(adate)= " & i
            gdBase.Execute sSQL, dbFailOnError
            alleWeekdays = False
        End If
    Next i
 
    If Not alleWeekdays Then
        sAuswertkr = ""
        For i = 2 To 7
            If Check1(i).Value = vbChecked Then
                If sAuswertkr = "" Then
                    sAuswertkr = gcWochentag(i - 1)
                Else
                    sAuswertkr = sAuswertkr & ", " & gcWochentag(i - 1)
                End If
            End If
        Next i
        If Check1(1).Value = vbChecked Then
            If sAuswertkr = "" Then
                sAuswertkr = gcWochentag(7)
            Else
                sAuswertkr = sAuswertkr & ", " & gcWochentag(7)
            End If
        End If
    End If
        
    sSQL = "Insert into zhead (auszr,auskr) values "
    sSQL = sSQL & " ( '" & sAuswertzr & "','" & sAuswertkr & "' )"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "zadateanz", gdBase
    sSQL = " select distinct adate into zadateanz from zeitzone "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select count(*)as anzTage from zadateanz "
    Set rszeitz = gdBase.OpenRecordset(sSQL)
    If Not rszeitz.RecordCount = 0 Then
        lanztage = rszeitz!anzTage
    End If
    rszeitz.Close
    
    Set rsInt = gdBase.OpenRecordset("zeitintv", dbOpenTable)
    If Not rsInt.EOF Then
        rsInt.MoveFirst
        Do While Not rsInt.EOF
        
            If Not IsNull(rsInt!timeint) Then
                timebis = TimeValue(rsInt!timeint)
            Else
                timebis = ""
            End If

            If Option1(0).Value Then 'alle 30min
            
                If timebis = "00:00:00" Then
                    timevon = "23:30:01"
                    timebis = "23:59:00"
                ElseIf Mid(timebis, 4, 2) = "30" Then
                    timevon = Left(timebis, 3) & "00:01"
                ElseIf Mid(timebis, 4, 2) = "00" Then
                    
                    timevon = Format(CInt(Left(timebis, 2)) - 1, "00") & ":30:01"
                End If
            Else 'alle Stunden
                If timebis = "00:00:00" Then
                    timevon = "23:00:01"
                    timebis = "23:59:00"
                ElseIf Mid(timebis, 4, 2) = "30" Then
                    timevon = Left(timebis, 3) & "00:01"
                ElseIf Mid(timebis, 4, 2) = "00" Then
                    
                    timevon = Format(CInt(Left(timebis, 2)) - 1, "00") & ":" & "00:01"
                End If
            End If
            
            timevon = "#" & timevon & "#"
            timebis = "#" & timebis & "#"
            
            sSQL = " select sum(preis) as preis1,sum(Menge)as menge1,sum(nse) as nse1,avg(nsp)as nsp1 from zeitzone "
            sSQL = sSQL & " where atime between " & timevon & " "
            sSQL = sSQL & "  and " & timebis & " "
            Set rszeitz = gdBase.OpenRecordset(sSQL)
            
            If Not rszeitz.EOF Then
                rsInt.Edit
                rsInt!Preis = rszeitz!preis1 / lanztage
                rsInt!Menge = rszeitz!menge1 / lanztage
                rsInt!PREISAbs = rszeitz!preis1
                rsInt!MengeAbs = rszeitz!menge1
                
                rsInt!nsp = rszeitz!nsp1 / lanztage
                rsInt!nse = rszeitz!nse1 / lanztage
                rsInt!nspABS = rszeitz!nsp1
                rsInt!nseABS = rszeitz!nse1
                
                
                rsInt.Update
            End If
            rszeitz.Close
            
            sSQL = "Select distinct Kasnum "
            sSQL = sSQL & "from zeitzone "
            Set rsKass = gdBase.OpenRecordset(sSQL)
            If Not rsKass.EOF Then
            
                rsKass.MoveFirst
                Do While Not rsKass.EOF
                
                    If Not IsNull(rsKass!kasnum) Then
                    
                        loeschNEW "zbonanz", gdBase
            
                        sSQL = " select distinct belegnr,adate into zbonanz from zeitzone "
                        sSQL = sSQL & " where atime between " & timevon & " "
                        sSQL = sSQL & "  and " & timebis & " "
                        sSQL = sSQL & "  and kasnum = " & rsKass!kasnum
                        sSQL = sSQL & " group by adate,belegnr "
                        gdBase.Execute sSQL, dbFailOnError
                        
                        sSQL = " select count(*)as kanz1 from zbonanz "
                        Set rszeitz = gdBase.OpenRecordset(sSQL)
                        
                        If Not rszeitz.EOF Then
                            rsInt.Edit
                            If lanztage > 0 Then
                                rsInt!kanz = rszeitz!kanz1 / lanztage
                                
                                If Not IsNull(rsInt!kanzabs) Then
                                    rsInt!kanzabs = rsInt!kanzabs + rszeitz!kanz1
                                Else
                                    rsInt!kanzabs = rszeitz!kanz1
                                End If
                            Else
                                rsInt!kanz = 0
                            End If
                            
                            rsInt.Update
                        End If
                        rszeitz.Close
                    End If
                    rsKass.MoveNext
                Loop
            End If
            rsKass.Close: Set rsKass = Nothing
            
            rsInt.MoveNext
        Loop
    End If
    rsInt.Close
    
    sSQL = "Delete from zeitintv where preis is null"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rszeitz = gdBase.OpenRecordset("zeitintv", dbOpenTable)
    If Not rszeitz.RecordCount = 0 Then
        erstelleDaten = True
    Else
        erstelleDaten = False
        Label3.ForeColor = vbRed
        Label3.Caption = "Es wurden keine Daten ermittelt."
        Label3.Refresh
    End If
    rszeitz.Close
    
'    loeschNEW "Kasstemp_" & srechnertab, gdBase
    
    Screen.MousePointer = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "erstelleDaten"
    Fehler.gsFehlertext = "Im Programmteil Zeitenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
'    Resume Next
    
End Function
