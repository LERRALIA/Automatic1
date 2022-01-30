VERSION 5.00
Begin VB.Form frmWKL58 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Kartenleser"
   ClientHeight    =   4875
   ClientLeft      =   1125
   ClientTop       =   945
   ClientWidth     =   9840
   Icon            =   "frmWKL58.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   4875
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Schlieﬂen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   5880
      TabIndex        =   25
      Top             =   3720
      Width           =   3855
   End
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
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "COM-Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3615
      Index           =   4
      Left            =   7800
      TabIndex        =   19
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   22
         Left            =   120
         TabIndex        =   29
         Top             =   2880
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   21
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   20
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   19
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   18
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   17
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   16
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   15
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Parit‰t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3615
      Index           =   3
      Left            =   5880
      TabIndex        =   15
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Keine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   14
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Ungerade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   13
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Gerade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   12
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Daten-Bits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3615
      Index           =   2
      Left            =   3960
      TabIndex        =   11
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   11
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   9
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Stop-Bits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3615
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "1,5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Baudrate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "19200"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "9600"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "4800"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "2400"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "1200"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmWKL58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
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
    Screen.MousePointer = 11
    Dim iCount As Integer
    Dim ctmp As String
    
    Select Case Index
        Case Is = 0
            'Setze Baudrate
            For iCount = 0 To 5
                If Option1(iCount).Value = True Then
                    ctmp = Option1(iCount).Caption
                    Exit For
                End If
            Next iCount
            gVerbindung.lBaudRate = Val(ctmp)
            
            'Setze Stop-Bits
            If Option1(6).Value = True Then
                gVerbindung.sStopBits = 1
            ElseIf Option1(7).Value = True Then
                gVerbindung.sStopBits = 1.5
            Else
                gVerbindung.sStopBits = 2
            End If
            
            'Setze Datenbits
            For iCount = 9 To 11
                If Option1(iCount).Value = True Then
                    ctmp = Option1(iCount).Caption
                    Exit For
                End If
            Next iCount
            gVerbindung.iDatenBits = Val(ctmp)
            
            'Setze Parit‰t
            If Option1(12).Value = True Then
                gVerbindung.cParitaet = "E"
            ElseIf Option1(13).Value = True Then
                gVerbindung.cParitaet = "O"
            Else
                gVerbindung.cParitaet = "N"
            End If
            
            'Setze ComPort
            For iCount = 15 To 22
                If Option1(iCount).Value = True Then
                    ctmp = Option1(iCount).Caption
                    Exit For
                End If
            Next iCount
            gVerbindung.iComPort = Val(ctmp)
            
            SchreibeDatenVerbindung
            
            Unload frmWKL58
            
        Case Is = 1
            Unload frmWKL58
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    For lcount = 0 To 5
        If Val(Option1(lcount).Caption) = gVerbindung.lBaudRate Then
            Option1(lcount).Value = True
            Exit For
        End If
    Next lcount
    
    If gVerbindung.sStopBits = 1 Then
        Option1(6).Value = True
    ElseIf gVerbindung.sStopBits = 1.5 Then
        Option1(7).Value = True
    Else
        Option1(8).Value = True
    End If
    
    For lcount = 9 To 11
        If Val(Option1(lcount).Caption) = gVerbindung.iDatenBits Then
            Option1(lcount).Value = True
            Exit For
        End If
    Next lcount
    
    If gVerbindung.cParitaet = "E" Then
        Option1(12).Value = True
    ElseIf gVerbindung.cParitaet = "O" Then
        Option1(13).Value = True
    Else
        Option1(14).Value = True
    End If
    
    For lcount = 15 To 22
        If Val(Option1(lcount).Caption) = gVerbindung.iComPort Then
            Option1(lcount).Value = True
            Exit For
        End If
    Next lcount
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub SchreibeDatenVerbindung()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As String
    Dim ctmp As String
    Dim cTmp2 As String
    
    ctmp = Format$(gVerbindung.lBaudRate, "#####0")
    ctmp = ctmp & ","
    ctmp = ctmp & gVerbindung.cParitaet
    ctmp = ctmp & ","
    ctmp = ctmp & Format$(gVerbindung.iDatenBits, "0")
    ctmp = ctmp & ","
    If gVerbindung.sStopBits = 1.5 Then
        cTmp2 = "1.5"
    Else
        cTmp2 = Format$(gVerbindung.sStopBits, "0")
    End If
    ctmp = ctmp & cTmp2
    
    gVerbindung.cSettings = ctmp
    
    iFileNr = FreeFile
    Open gcPfad & "VERBIND.TXT" For Binary As #iFileNr
    Close iFileNr
    Kill gcPfad & "VERBIND.TXT"
    
    iFileNr = FreeFile
    Open gcPfad & "VERBIND.TXT" For Binary As #iFileNr
    Put #iFileNr, 1, gVerbindung
    Close iFileNr
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenVerbindung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


