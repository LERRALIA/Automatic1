VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form dlgGutschein 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   2655
   ClientLeft      =   2760
   ClientTop       =   3450
   ClientWidth     =   7440
   ControlBox      =   0   'False
   Icon            =   "dlgGutschein.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Pfeil
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'Kein
      Height          =   8295
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
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
         Caption         =   "als Gutschein"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   1
         Left            =   4200
         TabIndex        =   1
         Top             =   1680
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "Geld auszahlen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Stornobetrag, Was nun?"
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
         TabIndex        =   3
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "dlgGutschein"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_R�ckgabewert As Integer

Public Property Get Back() As Integer
    Back = m_R�ckgabewert
End Property

Private Sub Command1_Click(Index As Integer)

    m_R�ckgabewert = 0
    
    Select Case Index
        Case 0
             m_R�ckgabewert = 1 'als Gutschein
        Case 1
             m_R�ckgabewert = 2 'Geld auszahlen
        
        
    End Select
    
    Unload dlgGutschein

End Sub
Private Sub Form_Load()

    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
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





