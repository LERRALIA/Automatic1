VERSION 5.00
Begin VB.Form frmWKL171
   Caption         =   " - DATEV Konten"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command0 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   4080
      Picture         =   "frmZEN174.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   40
      ToolTipText     =   "Kalender"
      Top             =   1080
      Width           =   480
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   3600
      Picture         =   "frmZEN174.frx":039F
      Style           =   1  'Grafisch
      TabIndex        =   39
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   3600
      Picture         =   "frmZEN174.frx":03EE
      Style           =   1  'Grafisch
      TabIndex        =   38
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datev Konten"
      Height          =   2295
      Left            =   5040
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton Command5 
         Caption         =   "Löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4320
         TabIndex        =   35
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "nur gewählte Filiale anzeigen"
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   6375
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         Style           =   2  'Dropdown-Liste
         TabIndex        =   31
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         MaxLength       =   4
         TabIndex        =   28
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   4320
         TabIndex        =   27
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   26
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6120
         TabIndex        =   17
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kontenbezeichnung:"
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
         Index           =   10
         Left            =   1200
         TabIndex        =   32
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Konto"
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
         Index           =   8
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Filiale:"
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
         Index           =   7
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Datev Konten"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   8160
      TabIndex        =   15
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filialkonten/Kostenstellen"
      Height          =   2055
      Left            =   0
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   9135
      Begin VB.CommandButton Command5 
         Caption         =   "Löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   4320
         TabIndex        =   36
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   24
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4320
         TabIndex        =   23
         Top             =   480
         Width           =   1695
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   6375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6120
         TabIndex        =   14
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Filiale:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Filialkonto"
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
         Index           =   3
         Left            =   1560
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kostenstelle"
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
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Filialkonten / Kostenstellen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   9960
      TabIndex        =   12
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox cboFil 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   10
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   9
      Left            =   1320
      Picture         =   "frmZEN174.frx":043D
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   10
      Left            =   1320
      Picture         =   "frmZEN174.frx":048C
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command0 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   1800
      Picture         =   "frmZEN174.frx":04DB
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Kalender"
      Top             =   1080
      Width           =   480
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   11280
      TabIndex        =   4
      Top             =   360
      Width           =   345
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5040
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Schließen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "bis:"
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
      Index           =   2
      Left            =   2400
      TabIndex        =   41
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Filiale:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "von:"
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
      TabIndex        =   9
      Top             =   960
      Width           =   1215
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
      Caption         =   "DATEV Konten"
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
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmZEN174"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check1_Click()
On Error GoTo LOKAL_ERROR

   ZeigeKonten
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 1        ' Kalender
            Text1(9).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 0        ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lDat            As Long
    
    Select Case Index
        Case 10
            If IsDate(Text1(9).Text) = False Then
                Text1(9).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(9).Text) = True Then
                    lDat = CLng(DateValue(Text1(9).Text))
                End If
                lDat = lDat + 1
                Text1(9).Text = Format(lDat, "DD.MM.YYYY")
            End If
        
        Case 9
            If IsDate(Text1(9).Text) = False Then
                Text1(9).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(9).Text) = True Then
                    lDat = CLng(DateValue(Text1(9).Text))
                End If
                lDat = lDat - 1
                Text1(9).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 1
            If IsDate(Text1(3).Text) = False Then
                Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(3).Text) = True Then
                    lDat = CLng(DateValue(Text1(3).Text))
                End If
                lDat = lDat + 1
                Text1(3).Text = Format(lDat, "DD.MM.YYYY")
            End If
        
        Case 0
            If IsDate(Text1(3).Text) = False Then
                Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(3).Text) = True Then
                    lDat = CLng(DateValue(Text1(3).Text))
                End If
                lDat = lDat - 1
                Text1(3).Text = Format(lDat, "DD.MM.YYYY")
            End If
        
       
    End Select
    
err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lBis            As Long
    Dim lVon            As Long
    Dim iFil            As Integer
    Dim sdat            As String
    Dim lDay            As Long

    Select Case Index
        Case 0
            Unload frmZEN174
        Case 1 'Export
        
            If cboFil.Text = "Filiale auswählen" Then
                anzeige "rot", "Bitte wählen Sie eine Filiale!", Label1(4)
                cboFil.SetFocus
                Exit Sub
            Else
                iFil = CInt(Left$(cboFil.Text, 3))
            End If
            
            Screen.MousePointer = 11
            
            lVon = CLng(DateValue(Text1(9).Text))
            lBis = CLng(DateValue(Text1(3).Text))
            
            loeschNEW "DATEVEXPORT", gdbMdb
            CreateTableT2 "DATEVEXPORT", gdbMdb
            
            For lDay = lVon To lBis
                anzeige "normal", Format(lDay, "DD.MM.YY"), Label1(4)
                EXPORT lDay, iFil
            Next lDay
            
            sdat = Format$(Text1(9).Text, "DDMM") & Format$(Text1(3).Text, "DDMM")
            
            Screen.MousePointer = 0
            
            ExportCSV iFil, sdat
            
        Case 2
            gsHelpstring = "DATEV Konten"
            frmZEN110.Show 1
        Case 3
            Frame1.Visible = True
            Frame2.Visible = False
            ZeigeFilKonten
            
        Case 4
            Frame1.Visible = False
            
        Case 5
            Frame2.Visible = True
            Frame1.Visible = False
            fülleDATEVALLG Combo3
            ZeigeKonten
            
        Case 6
            Frame2.Visible = False
        Case 7
        
            If Combo1.Text = "Filiale auswählen" Then
                anzeige "rot", "Bitte wählen Sie eine Filiale!", Label1(4)
                Combo1.SetFocus
                Exit Sub
            Else
                SpeicherFilKonten CInt(Left$(Combo1.Text, 3)), Text1(0).Text, Text1(1).Text
                ZeigeFilKonten
            End If
        Case 8
        
            If Combo2.Text = "Filiale auswählen" Then
                anzeige "rot", "Bitte wählen Sie eine Filiale!", Label1(4)
                Combo2.SetFocus
                Exit Sub
            Else
                SpeicherKonten CInt(Left$(Combo2.Text, 3)), Val(Text1(2).Text), Combo3.Text
                ZeigeKonten
            End If
        Case 9
            LoescheKonten
            ZeigeKonten
        Case 10
            LoescheFilKonten
            ZeigeFilKonten
    End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ExportCSV(iFil As Integer, sdat As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer

   
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    sSQL = " Select "
    sSQL = sSQL & " FILIALE  "
    sSQL = sSQL & ", FILBEZ  "
    sSQL = sSQL & ", ZEITRAUMVON "
    sSQL = sSQL & ", ZEITRAUMBIS "
    sSQL = sSQL & ", KOST  "
    sSQL = sSQL & ", FILKONTO  "
    sSQL = sSQL & ", KONTO  "
    sSQL = sSQL & ", KONTOBEZ  "
    sSQL = sSQL & ", BETRAG "
    sSQL = sSQL & " from DATEVEXPORT "
    
    Set rsrs = gdbMdb.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
'

        sAusgabedatname = "DATEV" & sdat & "F" & iFil & ".csv"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "BOX\" & sAusgabedatname
        cPfad = cPfad1 & "BOX"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = "FILIALE;FILBEZ;ZEITRAUMVON;ZEITRAUMBIS;KOST;FILKONTO;KONTO;KONTOBEZ;BETRAG" & Chr$(13) & Chr$(10)

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 8
                If Not IsNull(rsrs.Fields(i)) Then

                    If i > 0 Then
                        If i = 2 Then
                            cSatz = cSatz & ";" & Format(rsrs.Fields(i), "DDMM")
                        ElseIf i = 3 Then
                            cSatz = cSatz & ";" & Format(rsrs.Fields(i), "DDMM")
                        ElseIf i = 8 Then
                        
                            If rsrs.Fields(i) = 0 Then
                                cSatz = cSatz & ";"
                            Else
                                cSatz = cSatz & ";" & rsrs.Fields(i)
                            End If
                        Else
                            cSatz = cSatz & ";" & rsrs.Fields(i)
                        End If
                    Else
                        cSatz = rsrs.Fields(i)
                    End If
                Else
                    If i > 0 Then
                        cSatz = cSatz & ";"
                    Else
                        cSatz = ""
                    End If
                End If
            Next i
        
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("DATEVEXPORT", gdbMdb) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Zentrale Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmZEN172.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Zentrale Information:"
        End If
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportCSV"
        Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub LoescheKonten()
On Error GoTo LOKAL_ERROR

    Dim bFound      As Boolean
    Dim sSQL        As String
    Dim iFil        As Integer
    Dim cKontobez   As String
    Dim lcount      As Long
    
    bFound = False
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount
    
    If Not bFound Then
        anzeige "rot", "Bitte markieren Sie eine Zeile", Label1(4)
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    anzeige "Normal", "", Label1(4)
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            iFil = CInt(Left(List2.list(lcount), 3))
            cKontobez = Trim(Right(List2.list(lcount), 36))
            
            sSQL = "Delete * from DATEVKONTEN where FILIALE = " & iFil
            sSQL = sSQL & " and KONTOBEZ = '" & cKontobez & "'"
            gdbMdb.Execute sSQL, dbFailOnError
        End If
    Next
    
    
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheKonten"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LoescheFilKonten()
On Error GoTo LOKAL_ERROR

    Dim bFound      As Boolean
    Dim sSQL        As String
    Dim iFil        As Integer
    Dim lcount      As Long
    
    bFound = False
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount
    
    If Not bFound Then
        anzeige "rot", "Bitte markieren Sie eine Zeile", Label1(4)
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    anzeige "Normal", "", Label1(4)
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            iFil = CInt(Left(List1.list(lcount), 3))
            
            sSQL = "Delete * from KOST where FILIALE = " & iFil
            gdbMdb.Execute sSQL, dbFailOnError
        End If
    Next
    
    
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheFilKonten"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ZeigeFilKonten()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    Dim cLBSatz     As String
    
    List1.Clear
    
    cSQL = "Select * from KOST order by FILIALE "
    Set rsrs = gdbMdb.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALE) Then
                cFeld = rsrs!FILIALE
            Else
                cFeld = "0"
            End If
            cLBSatz = Space(3 - Len(cFeld)) & cFeld & " "
            
            
            If Not IsNull(rsrs!FILBEZ) Then
                cFeld = rsrs!FILBEZ
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            

            If Not IsNull(rsrs!KOST) Then
                cFeld = rsrs!KOST
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(11 - Len(cFeld))
            
            If Not IsNull(rsrs!FilKonto) Then
                cFeld = rsrs!FilKonto
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(4 - Len(cFeld))
            
            List1.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeFilKonten"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeKonten()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    Dim cLBSatz     As String
   
    List2.Clear
    
    cSQL = "Select * from DATEVKONTEN "
    If Check1.Value = vbChecked Then
        If Combo2.Text = "Filiale auswählen" Then
            anzeige "rot", "Bitte wählen Sie eine Filiale!", Label1(4)
            Combo2.SetFocus
            Exit Sub
        Else
            cSQL = cSQL & " where Filiale = " & CInt(Left$(Combo2.Text, 3))
        End If
    End If
    cSQL = cSQL & " order by FILIALE "
    
    Set rsrs = gdbMdb.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALE) Then
                cFeld = rsrs!FILIALE
            Else
                cFeld = "0"
            End If
            
            cLBSatz = Space(3 - Len(cFeld)) & cFeld & " "
            
            If Not IsNull(rsrs!FILBEZ) Then
                cFeld = rsrs!FILBEZ
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(20 - Len(cFeld))
            
            If Not IsNull(rsrs!Konto) Then
                cFeld = rsrs!Konto
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(5 - Len(cFeld))
            
            If Not IsNull(rsrs!KONTOBEZ) Then
                cFeld = rsrs!KONTOBEZ
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKonten"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherFilKonten(iFil As Integer, cKost As String, lFilKonto As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
   
    
    sSQL = "Delete * from KOST where FILIALE = " & iFil
    gdbMdb.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KOST (FILIALE,FILBEZ,KOST,FILKONTO) values ( "
    sSQL = sSQL & iFil
    sSQL = sSQL & " , '" & ermFilbez(CLng(iFil)) & "' "
    sSQL = sSQL & " , '" & cKost & "' "
    sSQL = sSQL & " , " & lFilKonto & " "
    sSQL = sSQL & " ) "
    gdbMdb.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeFilKonten"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherKonten(iFil As Integer, lKonto As Long, cKontobez As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
   
    
    sSQL = "Delete * from DATEVKONTEN where FILIALE = " & iFil
    sSQL = sSQL & " and KONTOBEZ = '" & cKontobez & "'"
    gdbMdb.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DATEVKONTEN (FILIALE,FILBEZ,KONTO,KONTOBEZ) values ( "
    sSQL = sSQL & iFil
    sSQL = sSQL & " , '" & ermFilbez(CLng(iFil)) & "' "
    sSQL = sSQL & " , " & lKonto & " "
    sSQL = sSQL & " , '" & cKontobez & "' "
    sSQL = sSQL & " ) "
    gdbMdb.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherKonten"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermKonto(iFil As Integer, cKontobez As String) As Long
    On Error GoTo LOKAL_ERROR
    
    ermKonto = 0
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    sSQL = "Select KONTO from DATEVKONTEN where FILIALE = " & iFil
    sSQL = sSQL & " and KONTOBEZ = '" & cKontobez & "'"
    Set rsrs = gdbMdb.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Konto) Then
            ermKonto = rsrs!Konto
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermKonto"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermKOST(iFil As Integer) As String
    On Error GoTo LOKAL_ERROR
    
    ermKOST = ""
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    sSQL = "Select KOST from KOST where FILIALE = " & iFil
    Set rsrs = gdbMdb.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!KOST) Then
            ermKOST = rsrs!KOST
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermKOST"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermFilKonto(iFil As Integer) As Long
    On Error GoTo LOKAL_ERROR
    
    ermFilKonto = 0
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    sSQL = "Select FilKonto from KOST where FILIALE = " & iFil
    Set rsrs = gdbMdb.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!FilKonto) Then
            ermFilKonto = rsrs!FilKonto
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermFilKonto"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub EXPORT(lTag As Long, iFil As Integer)
On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim dBetrag         As Double
    Dim dECBetrag       As Double
    Dim dLSBetrag       As Double
    Dim cKostenstelle   As String
    Dim lFilKonto       As Long
    Dim lKonto          As Long
    
    cKostenstelle = ermKOST(iFil)
    lFilKonto = ermFilKonto(iFil)
    
    
    
    'Umsatz 19%
    dBetrag = ermgesUmsatzMwstAusZumsatz(CStr(lTag), CStr(lTag), iFil, "V")
    lKonto = ermKonto(iFil, "Umsatz volle MwSt")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz volle MwSt", dBetrag
    
    'Umsatz 7%
    dBetrag = ermgesUmsatzMwstAusZumsatz(CStr(lTag), CStr(lTag), iFil, "E")
    lKonto = ermKonto(iFil, "Umsatz erm MwSt")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz erm MwSt", dBetrag
    
    'KK AE
    dBetrag = -1 * ermgesKK(CStr(lTag), CStr(lTag), iFil, "AE")
    lKonto = ermKonto(iFil, "Kreditkarten Amex")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten Amex", dBetrag
    
    'KK VI
    dBetrag = -1 * ermgesKK(CStr(lTag), CStr(lTag), iFil, "VI")
    lKonto = ermKonto(iFil, "Kreditkarten Visa")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten Visa", dBetrag
    
    'KK EU
    dBetrag = -1 * ermgesKK(CStr(lTag), CStr(lTag), iFil, "EU")
    lKonto = ermKonto(iFil, "Kreditkarten Eurocard")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten Eurocard", dBetrag
    
    'Kreditverkauf
    dBetrag = -1 * ermgesKREDAusZumsatz(CStr(lTag), CStr(lTag), iFil)
    lKonto = ermKonto(iFil, "Kreditverkauf")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditverkauf", dBetrag
    
    'Ausgaben
    dBetrag = -1 * ermgesAUSZAHLUNG(CStr(lTag), CStr(lTag), iFil, "AUSZAHLUNG")
    lKonto = ermKonto(iFil, "Ausgaben")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Ausgaben", dBetrag
    
    'Geldtransit zur Bank
    dBetrag = -1 * ermgesABSCHOPF(CStr(lTag), CStr(lTag), iFil)
    lKonto = ermKonto(iFil, "Geldtransit zur Bank")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Geldtransit zur Bank", dBetrag
    
    
    
    'Kassendifferenzen
    dBetrag = ermgesKassendiff(CStr(lTag), CStr(lTag), iFil)
    lKonto = ermKonto(iFil, "Kassendifferenzen")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kassendifferenzen", dBetrag
    
    
    'KK EC
    dECBetrag = ermgesKK(CStr(lTag), CStr(lTag), iFil, "EC")
    
    'Lastschriften
    dLSBetrag = ermgesLASTZAHLTE(CStr(lTag), CStr(lTag), iFil)
    
    dBetrag = dECBetrag + dLSBetrag
    
    lKonto = ermKonto(iFil, "Karte EC/ELV")
    
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Karte EC/ELV", dBetrag
    
    'Gutscheinlösung
    dBetrag = ermgesGUTZ(CStr(lTag), CStr(lTag), iFil)
    lKonto = ermKonto(iFil, "Gutscheine")
    InsertDATEVEXPORT iFil, Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Gutscheine", dBetrag
    
    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EXPORT"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InsertDATEVEXPORT(ifilnr As Integer, cTag As String, cKost As String, lFilKonto As Long _
, lKonto As Long, cKontobez As String, dBetrag As Double)
On Error GoTo LOKAL_ERROR

    If dBetrag = 0 Then
        Exit Sub
    End If

    Dim sSQL        As String
    
    sSQL = "Insert into DATEVEXPORT ( "
    sSQL = sSQL & " FILIALE  "
    sSQL = sSQL & ", FILBEZ  "
    sSQL = sSQL & ", ZEITRAUMVON "
    sSQL = sSQL & ", ZEITRAUMBIS "
    sSQL = sSQL & ", KOST  "
    sSQL = sSQL & ", FILKONTO  "
    
    sSQL = sSQL & ", KONTO  "
    sSQL = sSQL & ", KONTOBEZ  "
    sSQL = sSQL & ", BETRAG "
    sSQL = sSQL & " ) "
    sSQL = sSQL & " values ( "
    sSQL = sSQL & ifilnr
    sSQL = sSQL & ", '" & ermFilbez(CLng(ifilnr)) & "'"
    sSQL = sSQL & ", '" & cTag & "'"
    sSQL = sSQL & ", '" & cTag & "'"
    sSQL = sSQL & ", '" & cKost & "'"
    sSQL = sSQL & ", " & lFilKonto & "  "
    
    sSQL = sSQL & ", " & lKonto & "  "
    sSQL = sSQL & ", '" & cKontobez & "'"
    sSQL = sSQL & ", '" & dBetrag & "'  "
    sSQL = sSQL & " ) "
    gdbMdb.Execute sSQL, dbFailOnError
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertDATEVEXPORT"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    
    PositionierenZ174
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    füllefil1 cboFil
    füllefil1 Combo1
    füllefil1 Combo2
    
    If NewTableSuchenDBKombiTH("KOST", gdbMdb) = False Then
        CreateTableT2 "KOST", gdbMdb
    End If
    
    If NewTableSuchenDBKombiTH("DATEVKONTEN", gdbMdb) = False Then
        CreateTableT2 "DATEVKONTEN", gdbMdb
    End If
    
    If NewTableSuchenDBKombiTH("DATEVALLG", gdbMdb) = False Then
        CreateTableT2 "DATEVALLG", gdbMdb
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Umsatz volle MwSt') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Umsatz erm MwSt') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Kreditkarten Amex') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Kreditkarten Eurocard') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Kreditkarten Visa') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Karte EC/ELV') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Kreditverkauf') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Geldtransit zur Bank') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Ausgaben') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Kassendifferenzen') "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('Gutscheine') "
        gdbMdb.Execute sSQL, dbFailOnError
        
    End If
    
    Text1(9).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
   
   
    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub füllefil1(cbox As ComboBox)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cSatz As String
    Dim cFeld As String
    
    cbox.Clear
    cbox.AddItem "Filiale auswählen"
    cbox.Text = "Filiale auswählen"

    sSQL = "Select * from filialen order by filialnr"
    Set rsrs = gdbMdb.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FilialNr) Then
            
                cFeld = rsrs!FilialNr
                cSatz = cSatz & Space(3 - Len(cFeld)) & cFeld
                
                If Not IsNull(rsrs!Filialname) Then
                
                    cFeld = rsrs!Filialname
                    cSatz = cSatz & Space(2) & cFeld
                
                    cbox.AddItem cSatz
                    
                End If
            End If
            cSatz = ""
            cFeld = ""
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllefil1"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fülleDATEVALLG(cbox As ComboBox)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    cbox.Clear
    cbox.AddItem "bitte auswählen"
    cbox.Text = "bitte auswählen"

    sSQL = "Select * from DATEVALLG "
    Set rsrs = gdbMdb.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!KONTOBEZ) Then
                cbox.AddItem rsrs!KONTOBEZ
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllefil1"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenZ174()
On Error GoTo LOKAL_ERROR

    With Frame1
        .Height = 5535
        .Left = 5040
        .Top = 1560
        .Width = 6615
    End With
    
    With Frame2
        .Height = 5535
        .Left = 5040
        .Top = 1560
        .Width = 6615
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenZ174"
    Fehler.gsFehlertext = "Im Programmteil DATEV Konten ist ein Fehler aufgetreten."

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
