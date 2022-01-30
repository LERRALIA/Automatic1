VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL06 
   Caption         =   "Produktlinienbearbeitung"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "frmWKL06.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11760
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5520
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   9375
   End
   Begin sevCommand3.Command Command0 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   19
      Top             =   1440
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   9600
      TabIndex        =   18
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Size            =   14.25
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
      Caption         =   "Anzeigen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   8280
      MaxLength       =   6
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin sevCommand3.Command Command0 
      Height          =   375
      Index           =   54
      Left            =   11280
      TabIndex        =   14
      Top             =   2640
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   6
      Left            =   9600
      TabIndex        =   11
      Top             =   6840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Size            =   14.25
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   9600
      TabIndex        =   10
      Top             =   6300
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Size            =   14.25
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
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   9600
      TabIndex        =   9
      Top             =   5760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Size            =   14.25
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
      Caption         =   "Neue Linie"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   8
      Top             =   5220
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Size            =   14.25
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
      Caption         =   "Felder leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   7
      Top             =   4680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Size            =   14.25
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   120
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1440
      Width           =   975
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   0
      Top             =   7380
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Size            =   14.25
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
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Lieferant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   20
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Lieferant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   9600
      TabIndex        =   16
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Linienbezeichnung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   13
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Linie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblAnzeige 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   9135
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Produktlinienbearbeitung"
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
      Top             =   120
      Width           =   7575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmWKL06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 0
            Text1_KeyUp 3, vbKeyF2, 0
        Case Is = 54
            Text1_KeyUp 2, vbKeyF2, 0
    End Select

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Function pruef() As Boolean
    On Error GoTo LOKAL_ERROR
    
    pruef = False
    
    If Text1(0).Text = "" Then
        Exit Function
    End If
    
    If Not IsNumeric(Text1(0).Text) Then
        Exit Function
    End If
    
    If Not IsNumeric(Text1(3).Text) Then
        Exit Function
    End If
    
    If Text1(1).Text = "" Then
                
        Exit Function
    Else
        Text1(1).Text = SwapStr(Text1(1).Text, "'", " ")
        Text1(1).Text = SwapStr(Text1(1).Text, "*", " ")
        If Text1(1).Text = "" Then
                
            Exit Function
        End If
    End If
        
    pruef = True
            
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pruef"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function SchreibeDatenWKL06() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim rs1             As Recordset
    Dim clpz            As String
    Dim dFeld           As Double
    Dim cLinbezeich     As String
    Dim cLinr           As String
    
    SchreibeDatenWKL06 = False
    
    clpz = Trim$(Text1(0).Text)
    cLinbezeich = Trim(Text1(1).Text)
    cLinr = Trim$(Text1(3).Text)
    
    cSQL = "Select * from LISRT where linr = " & cLinr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        anzeigeNew "rot", "Diese Lieferantennumer ist unbekannt.", lblAnzeige
        Text1(3).SetFocus
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    cSQL = "Select * from LINBEZ where LPZ = " & clpz & " and LINR = " & cLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!LPZ = clpz
    rsrs!LINBEZEICH = cLinbezeich
    rsrs!linr = cLinr
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    SchreibeDatenWKL06 = True
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL06"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Select Case Index
    
        Case Is = 0     'Speichern
        
            If pruef Then
                If SchreibeDatenWKL06 Then
                    If IsNumeric(Text1(2).Text) Then
                        FuelleListbox2WKL06 Trim(Text1(2).Text)
                    Else
                        FuelleListbox2WKL06 "alle"
                    End If
                    InitDialogWKL06
                End If
            Else
                anzeigeNew "rot", "Bitte überprüfen Sie Ihre Eingaben!", lblAnzeige
            End If
        Case Is = 1     'Leeren
            InitDialogWKL06
            Text1(0).SetFocus
            anzeigeNew "normal", "", lblAnzeige
        Case Is = 2     'Beenden
            loeschNEW "LINTE", gdBase
            Unload frmWKL06
            
        Case Is = 3     'Beenden
            If IsNumeric(Text1(2).Text) Then
                FuelleListbox2WKL06 Trim(Text1(2).Text)
            Else
                FuelleListbox2WKL06 "alle"
            End If
        Case Is = 4
            InitDialogWKL06
            Text1(0).SetFocus
            anzeigeNew "normal", "", lblAnzeige
        Case Is = 5
            If Not NewTableSuchenDBKombi("LINTE", gdBase) Then
                FuelleListbox2WKL06 "alle"
            End If
            
            anzeigeNew "normal", "Druckvorschau wird erstellt...", lblAnzeige
            reportbildschirm "dWKL12a", "aWKL06"
            anzeigeNew "normal", "", lblAnzeige
            
        Case Is = 6
            If List2.ListIndex < 0 Then
                anzeigeNew "rot", "Bitte einen Eintrag in der Liste markieren!", lblAnzeige
                List2.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            MoveList2FelderWKL06
            LoescheAGN
            InitDialogWKL06
        
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheAGN()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cFeld As String
    Dim cLinr As String
    
    cFeld = Text1(0).Text
    cFeld = Trim$(cFeld)
    
    cLinr = Text1(3).Text
    cLinr = Trim$(cLinr)
    
    cSQL = "Delete from LINBEZ where lpz = " & cFeld
    cSQL = cSQL & " and LINR = " & cLinr
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
   
    
    If IsNumeric(Text1(2).Text) Then
        FuelleListbox2WKL06 Text1(2).Text
    
    Else
        FuelleListbox2WKL06 "alle"
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheAGN"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    WKL06Positionieren
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    

    InitDialogWKL06
    
    List1.Clear
    List2.Clear
    List1.AddItem "Linie Linienbezeichnung              Lieferantenbezeichnung              Linr"
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List2_Click()
On Error GoTo LOKAL_ERROR
    
    MoveList2FelderWKL06
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MoveList2FelderWKL06()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    Dim cFeld As String
    
    cLBSatz = List2.list(List2.ListIndex)
    
    cFeld = Mid(cLBSatz, 1, 6)
    cFeld = Trim$(cFeld)
    Text1(0).Text = cFeld
    
    cFeld = Mid(cLBSatz, 7, 30)
    cFeld = Trim$(cFeld)
    Text1(1).Text = cFeld
    
    cFeld = Right(cLBSatz, 6)
    cFeld = Trim$(cFeld)
    Text1(3).Text = cFeld
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveList2FelderWKL06"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    
    Select Case Index
        Case 0, 2, 3
            cValid = "1234567890" & Chr$(8)
        Case 1
            cValid = Chr$(KeyAscii)
    End Select
    
    If InStr(cValid, Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(Chr$(KeyAscii))
    End If
        

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleListbox2WKL06(sLief As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim dWert       As Double
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim lAnz        As Long
    Dim lcount      As Long
    Dim cwhere     As String
    
    If sLief = "alle" Then
        cwhere = " where linr <> 300200 "
    Else
        cwhere = " where linr <> 300200 and linr =  " & sLief
    End If
    
    anzeigeNew "normal", "Die Linien werden ermittelt...", lblAnzeige
    
    List2.Clear
    
    
    Me.Refresh
    
    loeschNEW "LinTE", gdBase
    CreateTable "LINTE", gdBase
    
    cSQL = "Insert into LINTE Select LINR,LPZ,LINBEZEICH , '' as LIEFBEZ from LINBEZ " & cwhere
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update LINTE inner join lisrt on LINTE.linr = lisrt.linr "
    cSQL = cSQL & " Set LINTE.liefbez = lisrt.liefbez "
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from LINTE " & cwhere & " order by linr,lpz "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        lcount = lAnz
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            
'            anzeigeNew "normal", lcount & " Linien werden angezeigt...", lblAnzeige
'            lcount = lcount - 1
            
            If Not IsNull(rsrs!LPZ) Then
                dWert = rsrs!LPZ
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "#####0")
            cLBSatz = Space(5 - Len(cFeld)) & cFeld & " "
            
            If Not IsNull(rsrs!LINBEZEICH) Then
                cFeld = rsrs!LINBEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space(31 - Len(cFeld))

            
            If Not IsNull(rsrs!LIEFBEZ) Then
                cFeld = rsrs!LIEFBEZ
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            
            If Not IsNull(rsrs!linr) Then
                dWert = rsrs!linr
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "#####0")
            cLBSatz = cLBSatz & Space(6 - Len(cFeld)) & cFeld

            
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lAnz = 0 Then
        anzeigeNew "rot", "Keine Produktlinien wurden ermittelt.", lblAnzeige
    ElseIf lAnz = 1 Then
        anzeigeNew "normal", lAnz & " Produktlinien wurde ermittelt.", lblAnzeige
    Else
        anzeigeNew "normal", lAnz & " Produktlinien wurden ermittelt.", lblAnzeige
    End If
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleListbox2WKL06"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InitDialogWKL06()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = ""
'    Text1(2).Text = ""
    Text1(3).Text = ""

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InitDialogWKL06"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
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
Private Sub WKL06Positionieren()
    On Error GoTo LOKAL_ERROR
    
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL06Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcount As Long
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        
        Select Case Index
            Case Is = 2
                gF2Prompt.bMultiple = False
                gF2Prompt.cFeld = "LINR"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
                
                If IsNumeric(Text1(2).Text) Then
                    FuelleListbox2WKL06 Trim(Text1(2).Text)
                End If
            
            Case Is = 3
                gF2Prompt.bMultiple = False
                gF2Prompt.cFeld = "LINR"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
                
            
        End Select
        Text1(Index).SetFocus
        
        
    End If
    
    
    If KeyCode = vbKeyReturn Then
        If IsNumeric(Text1(2).Text) Then
            FuelleListbox2WKL06 Trim(Text1(2).Text)
        Else
            FuelleListbox2WKL06 "alle"
        End If
    End If
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Produktlinien bearbeiten ist ein Fehler aufgetreten. "
    
End Sub
