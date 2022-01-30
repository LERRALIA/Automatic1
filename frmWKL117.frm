VERSION 5.00
Begin VB.Form frmWKL117 
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6960
      TabIndex        =   10
      Text            =   "0,01"
      Top             =   360
      Width           =   855
   End
   Begin sevCommand3.Command Command4 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Pay"
      Height          =   495
      Left            =   8760
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zahlung"
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin VB.TextBox ctrlAmountEF 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox ctrlUsePinCB 
         Caption         =   "PIN verlangen"
         Height          =   315
         Left            =   840
         Style           =   1  'Grafisch
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Zahlung vornehmen"
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Betrag in Cent:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin sevCommand3.Command Command2 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Client ID setzen"
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin sevCommand3.Command Command3 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Storno"
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5400
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "TA NR:"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "frmWKL117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lForcePIN As Long

Public Sub Command1_Click()

'  Dim lCents As Long
'  Dim lRet As Long
'  Dim sBLZ As String * 8000
'  Dim lBuffer As Long
'
'  Dim lerrCode As Long
'  Dim serrMeldung As String * 8000
'
'
'
'    lCents = Val(ctrlAmountEF.Text)
'    lRet = ELMEPay(lCents, lForcePIN)
'
'
'    If lRet = 0 Then
'        lRet = ELMEGetPrint(sBLZ, 8000)
'        If lRet = 0 Then
'            MsgBox sBLZ
'        Else
'            MsgBox lRet
'        End If
'    Else
'        MsgBox "Fehler ELMEPay: " & lRet
'        lRet = ELMEGetLastError(lerrCode, serrMeldung, 8000)
'        If lRet = 0 Then
'            MsgBox serrMeldung
'        Else
'            MsgBox lRet
'        End If
'    End If
    
    
    
    
    
    
    
End Sub
Private Sub Command2_Click()
Dim lRet As Long

lRet = ELMESettings(vbNullString, "111", vbNullString, vbNullString, -1, -1, -1, -1, vbNullString)
MsgBox lRet
End Sub

Private Sub Command3_Click()

Dim lCents As Long
  Dim lRet As Long
  Dim sBLZ As String * 8000
  Dim lBuffer As Long
  
  Dim lerrCode As Long
  Dim serrMeldung As String * 8000
  
  

    
    lRet = ELMEReversal(CLng(Text1.Text))
  
  
    If lRet = 0 Then
        lRet = ELMEGetPrint(sBLZ, 8000)
        If lRet = 0 Then
            MsgBox sBLZ
        Else
            MsgBox lRet
        End If
    Else
        MsgBox "Fehler ELMEPay: " & lRet
        lRet = ELMEGetLastError(lerrCode, serrMeldung, 8000)
        If lRet = 0 Then
            MsgBox serrMeldung
        Else
            MsgBox lRet
        End If
    End If
End Sub

'Private Sub Command4_Click()
'
'Dim sTest2 As String
'sTest2 = Init()
'
'Text2.Text = sTest2
'Text2.Refresh
'
'End Sub

Private Sub ctrlAmountEF_KeyPress(KeyAscii As Integer)
  Const Numbers$ = "0123456789"
  If InStr(Numbers, Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
  End If
End Sub

Private Sub ctrlUsePinCB_Click()
  If lForcePIN = 0 Then
    lForcePIN = 1
  Else
    lForcePIN = 0
  End If
End Sub


Private Sub Form_Load()
  lForcePIN = 0
  Set ifsfAmount = Text3
  
End Sub


