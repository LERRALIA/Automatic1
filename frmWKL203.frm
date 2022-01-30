VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL203 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Bela"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.OptionButton Option1 
      Caption         =   "Artikel von I - Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   6495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Artikel von A - H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   2160
      Width           =   6495
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   3
      Left            =   9480
      TabIndex        =   0
      Top             =   7800
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
      Caption         =   "Weiter"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Markieren Sie den gewünschten Bereich!"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   9015
   End
   Begin VB.Label lblanzeige 
      BackColor       =   &H00C0C000&
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
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   9255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Bela - Welche Sortimente möchten Sie übernehmen?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
        
        
        Case 3
        
            
            If Option1(0).Value = True Then
            
                sSQL = "Delete * from IMPORTPRI where bezeich like 'I*' "
                
                sSQL = sSQL & " or bezeich like 'J*'"
                sSQL = sSQL & " or bezeich like 'K*'"
                sSQL = sSQL & " or bezeich like 'L*'"
                sSQL = sSQL & " or bezeich like 'M*'"
                sSQL = sSQL & " or bezeich like 'N*'"
                sSQL = sSQL & " or bezeich like 'O*'"
                sSQL = sSQL & " or bezeich like 'P*'"
                sSQL = sSQL & " or bezeich like 'Q*'"
                sSQL = sSQL & " or bezeich like 'R*'"
                
                sSQL = sSQL & " or bezeich like 'S*'"
                sSQL = sSQL & " or bezeich like 'T*'"
                sSQL = sSQL & " or bezeich like 'U*'"
                sSQL = sSQL & " or bezeich like 'V*'"
                sSQL = sSQL & " or bezeich like 'W*'"
                sSQL = sSQL & " or bezeich like 'X*'"
                sSQL = sSQL & " or bezeich like 'Y*'"
                sSQL = sSQL & " or bezeich like 'Z*'"
                gdBase.Execute sSQL, dbFailOnError
            
                
            
            
            Else
            
                sSQL = "Delete * from IMPORTPRI where bezeich like 'A*' "
                sSQL = sSQL & " or bezeich like 'B*'"
                sSQL = sSQL & " or bezeich like 'C*'"
                sSQL = sSQL & " or bezeich like 'D*'"
                sSQL = sSQL & " or bezeich like 'E*'"
                sSQL = sSQL & " or bezeich like 'F*'"
                sSQL = sSQL & " or bezeich like 'G*'"
                sSQL = sSQL & " or bezeich like 'H*'"
                sSQL = sSQL & " or bezeich like ' *'"
                sSQL = sSQL & " or bezeich like '+*'"
                sSQL = sSQL & " or bezeich like '-*'"
                sSQL = sSQL & " or bezeich like '_*'"
                sSQL = sSQL & " or bezeich like '0*'"
                sSQL = sSQL & " or bezeich like '1*'"
                sSQL = sSQL & " or bezeich like '2*'"
                sSQL = sSQL & " or bezeich like '3*'"
                sSQL = sSQL & " or bezeich like '4*'"
                sSQL = sSQL & " or bezeich like '5*'"
                sSQL = sSQL & " or bezeich like '6*'"
                sSQL = sSQL & " or bezeich like '7*'"
                sSQL = sSQL & " or bezeich like '8*'"
                sSQL = sSQL & " or bezeich like '9*'"
                gdBase.Execute sSQL, dbFailOnError
            
                
            
            
            
            
            End If
        
        
        
        
        
        
            Unload frmWKL203
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Bela Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    anzeige "normal", "", lblanzeige
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bela Warengruppen ist ein Fehler aufgetreten."
    
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




