VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL180 
   Caption         =   "Auswahl einer Gruppe"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL180.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   1
      Left            =   7320
      TabIndex        =   5
      Top             =   7800
      Width           =   2055
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
      Caption         =   "Abbrechen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame5"
      Height          =   6855
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   11775
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   2
         Left            =   6840
         TabIndex        =   8
         Top             =   5160
         Width           =   2295
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
         Caption         =   "Hinzufügen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Top             =   5160
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   0
         Left            =   240
         MaxLength       =   6
         TabIndex        =   6
         Top             =   5160
         Width           =   1935
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   8775
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   8775
      End
      Begin VB.Label Label1 
         Caption         =   "Gruppenbezeichnung"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   10
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "GruppenNr"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   4800
         Width           =   1935
      End
   End
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   0
      Left            =   9480
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
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
      Caption         =   "weiter"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   7920
      Width           =   6855
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Wählen Sie eine Gruppe"
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
      TabIndex        =   1
      Top             =   0
      Width           =   9495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0 'auswählen
        
            If List2.ListIndex < 0 Then
                anzeige "rot", "Bitte einen Eintrag auswählen!", Label1(2)
                Exit Sub
            End If
            gsGRUPPENNR = Trim(Left(List2.list(List2.ListIndex), 6))
            Unload frmWKL180
        Case 1
            Unload frmWKL180
        Case 2
            hinzufuegen_gruppe Text1(0).Text, Text1(1).Text
            zeige_Gruppen
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Gruppen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren
    
    Modul6.Skalieren Me, True, True:
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    zeige_Gruppen
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Gruppen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub zeige_Gruppen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim cLBSatz     As String
    Dim cGruppenNr  As String
    Dim cGruppenbez As String
    
    Screen.MousePointer = 11
    
    List1.Clear
    List2.Clear
    List1.AddItem "GruppenNr Gruppenbezeichnung"
    
    sSQL = "Select GruppenNr, Gruppenbez from Gruppe group by gruppenNr, Gruppenbez"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!Gruppennr) Then
                cGruppenNr = rsrs!Gruppennr
            End If
            
            If Not IsNull(rsrs!Gruppenbez) Then
                cGruppenbez = rsrs!Gruppenbez
            End If
            
            cLBSatz = cGruppenNr & Space$(8 - Len(cGruppenNr))
            cLBSatz = cLBSatz & cGruppenbez
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Gruppen"
    Fehler.gsFehlertext = "Im Programmteil Gruppen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub hinzufuegen_gruppe(sGruppenNr As String, sGruppenbez As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim iRet        As Integer
    Dim rsrs        As Recordset
    
    If sGruppenNr = "" Then
        Exit Sub
    Else
        If IsNumeric(sGruppenNr) = False Then
            Exit Sub
        End If
    End If

'    If sGruppenbez = "" Then
'        Exit Sub
'    End If

    sSQL = "Select * from Gruppe where GruppenNr = " & sGruppenNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        iRet = MsgBox("Diese Gruppennummer ist schon vorhanden, überscheiben?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    sSQL = "Delete from Gruppe where GruppenNr = " & sGruppenNr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Gruppe (GruppenNr,Gruppenbez) values (" & sGruppenNr & ",'" & sGruppenbez & "')"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "hinzufuegen_gruppe"
    Fehler.gsFehlertext = "Im Programmteil Gruppen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    With Frame5
        .Height = 6735
        .Left = 0
        .Top = 840
        .Width = 11775
        .BorderStyle = 0
        
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Gruppen ist ein Fehler aufgetreten."
    
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

