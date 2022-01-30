VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKLar 
   BackColor       =   &H00C0C000&
   Caption         =   "Bondruck Test"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   -180
   ClientWidth     =   5520
   Icon            =   "frmWKLar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5520
   StartUpPosition =   1  'Fenstermitte
   Begin sevCommand3.Command cmd1 
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
   Begin sevCommand3.Command cmd1 
      Height          =   615
      Index           =   0
      Left            =   2880
      TabIndex        =   1
      Top             =   2880
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
   Begin VB.ListBox Listx 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   3
      Left            =   7440
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   7440
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   7440
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmWKLar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            Unload frmWKLar
        Case 1
            DruckeZweitBonAusListe Listx, False
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEnd_Click"
    Fehler.gsFehlertext = "Im Programmteil Bon Test ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Form_Load()
   On Error GoTo LOKAL_ERROR
   
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing

    frmWKLar.Top = Screen.Height / 2 - frmWKLar.Height / 2
    frmWKLar.Left = Screen.Width / 2 - frmWKLar.Width / 2
    
    If giBonusNr = 0 Then
        TestVariante1
    ElseIf giBonusNr = 1 Then
        TestVariante2
    ElseIf giBonusNr = 2 Then
        TestVariante3
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bon Test ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TestVariante1()
    On Error GoTo LOKAL_ERROR
    
    Dim sEinText As String
    Dim sTeilString As String
    Dim sWort As String
    Dim iZeilenLen As Integer
    Dim cLBSatz As String
    iZeilenLen = 32
    
    Listx.Clear
    
    sEinText = Trim(gsTextVor) & " 145 " & Trim(gsTextNach) & " "
    
    If Len(sEinText) > iZeilenLen Then
        Do While Len(sEinText) > 0
            sWort = Mid(sEinText, 1, InStr(1, sEinText, " ") - 1)
            
            If Len(cLBSatz & sWort & Space(1)) > iZeilenLen Then
                Listx.AddItem cLBSatz
                cLBSatz = ""
            End If
            
            cLBSatz = cLBSatz & sWort & Space(1)
            sEinText = Mid(sEinText, Len(sWort) + 2, Len(sEinText) - Len(sWort) + 1)
        Loop
        Listx.AddItem cLBSatz
    Else
        cLBSatz = sEinText
        Listx.AddItem cLBSatz
    End If
    
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TestVariante1"
    Fehler.gsFehlertext = "Im Programmteil Bon Test ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TestVariante2()
    On Error GoTo LOKAL_ERROR
    
    Dim sEinText As String
    Dim sTeilString As String
    Dim sWort As String
    Dim iZeilenLen As Integer
    Dim cLBSatz As String
    iZeilenLen = 32
    
    Listx.Clear
    
    sEinText = Trim(gsTextVor) & " 5,24 " & Trim(gsWWZeichen) & " " & Trim(gsTextNach) & " "
    
    If Len(sEinText) > iZeilenLen Then
        Do While Len(sEinText) > 0
            sWort = Mid(sEinText, 1, InStr(1, sEinText, " ") - 1)
            
            If Len(cLBSatz & sWort & Space(1)) > iZeilenLen Then
                Listx.AddItem cLBSatz
                cLBSatz = ""
            End If
            
            cLBSatz = cLBSatz & sWort & Space(1)
            sEinText = Mid(sEinText, Len(sWort) + 2, Len(sEinText) - Len(sWort) + 1)
        Loop
        Listx.AddItem cLBSatz
    Else
        cLBSatz = sEinText
        Listx.AddItem cLBSatz
    End If
    
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TestVariante1"
    Fehler.gsFehlertext = "Im Programmteil Bon Test ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TestVariante3()
    On Error GoTo LOKAL_ERROR
    
    Dim sEinText As String
    Dim sTeilString As String
    Dim sWort As String
    Dim iZeilenLen As Integer
    Dim cLBSatz As String
    iZeilenLen = 32
    
    Listx.Clear
    
    sEinText = gsSpezBontextU
    
    sEinText = Space$((iZeilenLen - Len(sEinText)) / 2) & sEinText
    Listx.AddItem sEinText
    
    
    Listx.AddItem "________________________________"
    
    
    cLBSatz = ""
    sWort = ""
    
    sEinText = gsSpezBontext & " "
    
    If Len(sEinText) > iZeilenLen Then
        Do While Len(sEinText) > 0
            sWort = Mid(sEinText, 1, InStr(1, sEinText, " ") - 1)
            
            If Len(cLBSatz & sWort & Space(1)) > iZeilenLen Then
                cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
                Listx.AddItem cLBSatz
                cLBSatz = ""
            End If
            
            cLBSatz = cLBSatz & sWort & Space(1)
            sEinText = Mid(sEinText, Len(sWort) + 2, Len(sEinText) - Len(sWort) + 1)
        Loop
        cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
        Listx.AddItem cLBSatz
    Else
        cLBSatz = sEinText
        
        cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
        Listx.AddItem cLBSatz
    End If
    
    
    
    Listx.AddItem Space(32)
    cLBSatz = ""
    sWort = ""
    
    sEinText = gsSpezBontext2 & " "
    
    If Len(sEinText) > iZeilenLen Then
        Do While Len(sEinText) > 0
            sWort = Mid(sEinText, 1, InStr(1, sEinText, " ") - 1)
            
            If Len(cLBSatz & sWort & Space(1)) > iZeilenLen Then
                cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
                Listx.AddItem cLBSatz
                cLBSatz = ""
            End If
            
            cLBSatz = cLBSatz & sWort & Space(1)
            sEinText = Mid(sEinText, Len(sWort) + 2, Len(sEinText) - Len(sWort) + 1)
        Loop
        cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
        Listx.AddItem cLBSatz
    Else
        cLBSatz = sEinText
        
        cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
        Listx.AddItem cLBSatz
    End If
    
    Listx.AddItem Space(32)
    cLBSatz = ""
    sWort = ""
    
    sEinText = gsSpezBontext3 & " "
    
    If Len(sEinText) > iZeilenLen Then
        Do While Len(sEinText) > 0
            sWort = Mid(sEinText, 1, InStr(1, sEinText, " ") - 1)
            
            If Len(cLBSatz & sWort & Space(1)) > iZeilenLen Then
                cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
                Listx.AddItem cLBSatz
                cLBSatz = ""
            End If
            
            cLBSatz = cLBSatz & sWort & Space(1)
            sEinText = Mid(sEinText, Len(sWort) + 2, Len(sEinText) - Len(sWort) + 1)
        Loop
        cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
        Listx.AddItem cLBSatz
    Else
        cLBSatz = sEinText
        
        cLBSatz = Space$((iZeilenLen - Len(cLBSatz)) / 2) & cLBSatz
        Listx.AddItem cLBSatz
    End If
    
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    Listx.AddItem Space(32)
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TestVariante3"
    Fehler.gsFehlertext = "Im Programmteil Bon Test ist ein Fehler aufgetreten."
    
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

