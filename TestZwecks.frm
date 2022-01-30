VERSION 5.00
Begin VB.Form TestZwecks 
   Caption         =   "TestZwecks"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "SFTP Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "TestZwecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sHosti  As String
Dim sUseri As String
Dim sPassi As String

Public WithEvents rfConnection As cConnection
Attribute rfConnection.VB_VarHelpID = -1
 
 
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR:
 
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = ""
    
    Fehlermeldung1
End Sub
