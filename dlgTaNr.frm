VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form dlgTaNr 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   3735
   ClientLeft      =   2760
   ClientTop       =   3450
   ClientWidth     =   7065
   ControlBox      =   0   'False
   Icon            =   "dlgTaNr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Pfeil
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'Kein
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1920
         TabIndex        =   16
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'Kein
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   6735
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "1"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   1
            Left            =   600
            TabIndex        =   6
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "2"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   2
            Left            =   1200
            TabIndex        =   7
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "3"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   3
            Left            =   1800
            TabIndex        =   8
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "4"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   4
            Left            =   2400
            TabIndex        =   9
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "5"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   5
            Left            =   3000
            TabIndex        =   10
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "6"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   6
            Left            =   3600
            TabIndex        =   11
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "7"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   7
            Left            =   4200
            TabIndex        =   12
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "8"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   8
            Left            =   4800
            TabIndex        =   13
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "9"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   9
            Left            =   5400
            TabIndex        =   14
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "0"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   855
            Index           =   14
            Left            =   6120
            TabIndex        =   15
            Top             =   0
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1023
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "C"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   3
         Left            =   4680
         TabIndex        =   3
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "OK"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Zentriert
         BackColor       =   &H000000C0&
         Caption         =   "Stornierung einer Kartenzahlung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Zentriert
         BackColor       =   &H000000C0&
         Caption         =   "Geben Sie bitte die BNr.:(steht auf dem Bon) ein! Bedienen Sie dann das Kartenterminal"""
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   6855
      End
   End
End
Attribute VB_Name = "dlgTaNr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_Rückgabewert As Long

Public Property Get Back() As Long
    Back = m_Rückgabewert
End Property
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Text1.Text = "" Then
        Text1.SetFocus
    Else
        m_Rückgabewert = Text1.Text
        Unload dlgTaNr
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
        
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If Index = 14 Then
        Text1.Text = ""
    Else
        Text1.Text = Text1.Text & SSCommand2(Index).Caption
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
