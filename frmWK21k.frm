VERSION 5.00
Begin VB.Form frmWK21k 
   BackColor       =   &H000000C0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   3915
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmWK21k.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.Timer Timer1 
         Left            =   3720
         Top             =   2280
      End
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Winkiss beenden"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   3000
         Width           =   5655
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Diese Meldung bleibt 60 Sekunden stehen"
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
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Winkiss lässt sich erst nach der Übertragung der Dateien starten."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Sie haben noch nicht die Daten übertragen. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmWK21k"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR

    AbmeldungDabaNew
    End
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Statistikhinweis ist ein Fehler aufgetreten."
    
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
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Timer1.Interval = 60000
    Timer1.Enabled = True
    
  Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Statistikhinweis ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer1_Timer()
On Error GoTo LOKAL_ERROR

    Unload frmWK21k
    Timer1.Enabled = False
    
 Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil Statistikhinweis ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
