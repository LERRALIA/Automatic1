VERSION 5.00
Begin VB.Form ChooseFile 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox gewaehltePfad 
      Enabled         =   0   'False
      Height          =   405
      Left            =   480
      TabIndex        =   3
      Top             =   680
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "abbrechen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "wählen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
End
Attribute VB_Name = "ChooseFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

gbDsFinvkPfad = ""
Unload Me

End Sub

Private Sub Command2_Click()

    If Trim(gewaehltePfad.Text) = "" Then
     MsgBox ("Bitte erstmal Pfad wählen ! ! !")
     
     Else
     
      gbDsFinvkPfad = gewaehltePfad.Text
      gbTSEExportPfad = gewaehltePfad.Text
      
      Unload Me
      
     
    End If

End Sub

Private Sub Dir1_Change()
gewaehltePfad.Text = Dir1.Path & "\"
End Sub

Private Sub Drive1_Change()
    
    Dir1.Path = Drive1.Drive
    
End Sub

