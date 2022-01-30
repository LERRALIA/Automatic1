VERSION 5.00
Begin VB.Form frmWKLah 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Tagesdatum umrechnen"
   ClientHeight    =   2460
   ClientLeft      =   3120
   ClientTop       =   1815
   ClientWidth     =   3930
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   2460
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Schlieﬂen"
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Leeren"
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Berechnen"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   120
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox MaskEdBox1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "01.01.1900"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "="
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "="
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Jahrhunderttag:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Datum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmWKLah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LeereDialogWKLah()
    
   
    Text1.Text = ""
    Label1(4).Caption = ""
    Label1(5).Caption = ""

End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Dim cDatum1 As String
    Dim cDatum2 As String
    Dim lWert1 As Long
    Dim lWert2 As Long
    
    Select Case Index
        Case Is = 0
            cDatum1 = Text1.Text
            If Trim$(Text1.Text) <> "" Then
                lWert1 = Val(Text1.Text)
            Else
                
                lWert1 = Fix(Now)
                Text1.Text = Trim$(Str$(lWert1))
            End If
            
            If cDatum1 <> "__.__.____" Then
                If IsDate(cDatum1) Then
                    lWert2 = DateValue(cDatum1)
                    Label1(4).Caption = Format$(lWert2, "####0")
                End If
            End If
            
            cDatum2 = Format$(lWert1, "DD.MM.YYYY")
            Label1(5).Caption = cDatum2
            
            
        Case Is = 1
            Unload frmWKLah
        
        Case Is = 2
            LeereDialogWKLah
            
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    MsgBox "frmWKLah.Command1_Click: " & Err.Number & " / " & Err.Description
End Sub


Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    frmWKLah.Top = Screen.Height / 2 - frmWKLah.Height / 2
    frmWKLah.Left = Screen.Width / 2 - frmWKLah.Width / 2
    
    LeereDialogWKLah
    
    
Exit Sub
LOKAL_ERROR:
    MsgBox "frmWKLah.Form_Load: " & Err.Number & " / " & Err.Description
End Sub


