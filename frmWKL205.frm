VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL205 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Staffelpreise"
   ClientHeight    =   7560
   ClientLeft      =   2055
   ClientTop       =   2865
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   7560
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   24
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   23
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   22
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   21
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   240
      MaxLength       =   35
      TabIndex        =   8
      Top             =   2880
      Width           =   3015
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      Style           =   2  'Dropdown-Liste
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   7335
   End
   Begin sevCommand3.Command Command2 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   6960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Entferne Artikel aus dieser Gruppe"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin sevCommand3.Command Command2 
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Bearbeiten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Neu"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   11
      Top             =   2800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin sevCommand3.Command Command2 
      Height          =   375
      Index           =   6
      Left            =   6240
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin sevCommand3.Command Command2 
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   16
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin sevCommand3.Command Command2 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Artikel zur ausgewählten Gruppe hinzufügen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   375
      Index           =   7
      Left            =   6240
      TabIndex        =   20
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Etikett"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "befindet sich in dieser Gruppe"
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
      Height          =   735
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "StaffelNr"
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
      Left            =   6720
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staffelgruppen-Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staffelpreis-Gruppe"
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
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preis(KVK)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Artnr"
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
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staffelpreise"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmWKL205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Combo3_Click()
On Error GoTo LOKAL_ERROR

    Dim sStaffNr As String
    
    sStaffNr = Right(Combo3.Text, 4)

    ZeigeArtikel_StaffelGruppe Val(sStaffNr), List2
    
    UnsichtbarSG False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_Click"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeArtikel_StaffelGruppe(iStaffNr As Integer, Listx As ListBox)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim cLBSatz As String
    Dim sFeld As String
    
    Listx.Clear

    cSQL = "Select * from STAFFEL_KVK_ARTIKEL where StaffNr = " & iStaffNr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            cLBSatz = ""
            sFeld = ""
        
            If Not IsNull(rsrs!artnr) Then
                sFeld = rsrs!artnr
            End If
            
            cLBSatz = cLBSatz & Space$(6 - Len(sFeld))
            cLBSatz = cLBSatz & sFeld & Space$(2)
            
            cSQL = "Select * from ARTIKEL where Artnr = " & sFeld & " "
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst
                
                If Not IsNull(rsArt!KVKPR1) Then
                    sFeld = Format(rsArt!KVKPR1, "######0.00")
                End If
                
                cLBSatz = cLBSatz & Space$(10 - Len(sFeld))
                cLBSatz = cLBSatz & sFeld & Space$(2)
                
                If Not IsNull(rsArt!BEZEICH) Then
                    sFeld = rsArt!BEZEICH
                End If
                
                cLBSatz = cLBSatz & sFeld & Space$(37 - Len(sFeld))
                
            End If
            rsArt.Close
            
            Listx.AddItem cLBSatz
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeArtikel_StaffelGruppe"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub




Private Sub Command5_Click(Index As Integer)

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

Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    Select Case Index
        Case 0      'Schließen
            Unload frmWKL205
        Case 1     'Lösche Artikel aus Gruppe
            deleteStaffelpreis List2
            UnsichtbarSG False
        Case 2 'Artikel Hinzufügen zur Gruppe
            Artikel_zur_Gruppe_hinzufuegen
            UnsichtbarSG False
        Case 3 'Bearbeiten
            ZeigeStaffelGruppe_Zum_Bearbeiten
        Case 4 'neu Staffelgruppe
            UnsichtbarSG True
            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(2).Text = ""
            
            Text1(2).SetFocus
            Label1(2).Caption = "0"
            
        Case 5 'Speichern
            SpeicherStaffelGruppe
            UnsichtbarSG False
        Case 6 'Lösche Gruppe
        
            loescheKompletteGruppe
            fülleStaffelGruppen Combo3
            UnsichtbarSG False
        Case 7 'drucke Etikett
            DruckeStaffelEtikett
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeStaffelGruppe_Zum_Bearbeiten()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    Dim sStaffNr    As String
   
    sStaffNr = Right(Combo3.Text, 4)
    
    If Val(sStaffNr) = 0 Then
        MsgBox "Bitte wählen Sie eine Staffelgruppe aus!", vbInformation, "Winkiss Hinweis:"
        Label1(2).Caption = "0"
        Exit Sub
    End If
    
    Label1(2).Caption = Val(sStaffNr)
    
    sSQL = "Select * from STAFFEL_KVK_GRUPPE where Staffnr = " & Val(sStaffNr)
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!StaffName) Then
            Text1(2).Text = rsrs!StaffName
        End If
        
        If Not IsNull(rsrs!Menge) Then
            Text1(0).Text = rsrs!Menge
        End If
        
        If Not IsNull(rsrs!KVKPR1) Then
            Text1(1).Text = Format(rsrs!KVKPR1, "######0.00")
        End If
        
        
        
        
        
        If Not IsNull(rsrs!Menge2) Then
            Text1(3).Text = rsrs!Menge2
        End If
        
        If Not IsNull(rsrs!KVKPR2) Then
            Text1(4).Text = Format(rsrs!KVKPR2, "######0.00")
        End If
        
        
        
        If Not IsNull(rsrs!Menge3) Then
            Text1(6).Text = rsrs!Menge3
        End If
        
        If Not IsNull(rsrs!KVKPR3) Then
            Text1(5).Text = Format(rsrs!KVKPR3, "######0.00")
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    UnsichtbarSG True
    
    Text1(2).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeStaffelGruppe_Zum_Bearbeiten"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeStaffelEtikett()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim sStaffNr    As String
   
    sStaffNr = Right(Combo3.Text, 4)
    
    If Val(sStaffNr) = 0 Then
        MsgBox "Bitte wählen Sie eine Staffelgruppe aus!", vbInformation, "Winkiss Hinweis:"
'        Label1(2).Caption = "0"
        Exit Sub
    End If
    
    loeschNEW "VORTEILSETI", gdBase
    
    sSQL = "Create Table VORTEILSETI ("
    sSQL = sSQL & " STAFFUEBERText1 Text(50)"
    sSQL = sSQL & ", STAFFMENGEText1 Text(50)"
    sSQL = sSQL & ", STAFFMENGE long"
    sSQL = sSQL & ", STAFFPREIS Double"
    sSQL = sSQL & ", STAFFPREISText1 Text(50)"
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into VORTEILSETI ("
    sSQL = sSQL & " STAFFUEBERText1 "
    sSQL = sSQL & ", STAFFMENGEText1 "
    sSQL = sSQL & ", STAFFMENGE "
    sSQL = sSQL & ", STAFFPREIS "
    sSQL = sSQL & ", STAFFPREISText1 "
    sSQL = sSQL & ") values "
    sSQL = sSQL & " ( 'Vorteilspreis' "
    sSQL = sSQL & ", 'ab' "
    sSQL = sSQL & ", " & Text1(0).Text & " "
    sSQL = sSQL & ", '" & Text1(1).Text & "' "
    sSQL = sSQL & ", 'Stück' "
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError
    

    
   
    
    reportbildschirmToPrinterETI "STAFFETI_1", gcEtikettenDrucker, True
    
    
    setzedrucker gcListenDrucker
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeStaffelEtikett"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherStaffelGruppe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lStaffNr As Long
    Dim sStaffName As String
    
    Dim dKVK As Double
    Dim iMenge As Integer
    
    Dim dKVK2 As Double
    Dim iMenge2 As Integer
    
    Dim dKVK3 As Double
    Dim iMenge3 As Integer
    
    lStaffNr = Val(Label1(2).Caption)
    
    If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Then
    
        Exit Sub
    End If
    
    sStaffName = Text1(2).Text
    
    dKVK = Text1(1).Text
    iMenge = Text1(0).Text
    
    If Text1(4).Text = "" Then Text1(4).Text = "0"
    If Text1(3).Text = "" Then Text1(3).Text = "0"
    If Text1(5).Text = "" Then Text1(5).Text = "0"
    If Text1(6).Text = "" Then Text1(6).Text = "0"
    
    dKVK2 = Text1(4).Text
    iMenge2 = Text1(3).Text
    
    dKVK3 = Text1(5).Text
    iMenge3 = Text1(6).Text
    
    If lStaffNr = 0 Then
            
        lStaffNr = ermMaxStaffNr
        Label1(2).Caption = lStaffNr
        
        'sicherheits löschen
        sSQL = "Delete from Staffel_KVK_Artikel where Staffnr = " & lStaffNr
        gdBase.Execute sSQL, dbFailOnError
        
        
    
    End If
            
            
    sSQL = "Delete from Staffel_KVK_Gruppe where Staffnr = " & lStaffNr
    gdBase.Execute sSQL, dbFailOnError


    sSQL = "Insert into STAFFEL_KVK_GRUPPE (STAFFNR,STAFFNAME,KVKPR1,MENGE,KVKPR2,MENGE2,KVKPR3,MENGE3)"
    sSQL = sSQL & " values "
    sSQL = sSQL & " ( " & lStaffNr & " "
    sSQL = sSQL & " , '" & sStaffName & "' "
    sSQL = sSQL & " , '" & dKVK & "' "
    sSQL = sSQL & " , " & iMenge & " "
    
    sSQL = sSQL & " , '" & dKVK2 & "' "
    sSQL = sSQL & " , " & iMenge2 & " "
    
    sSQL = sSQL & " , '" & dKVK3 & "' "
    sSQL = sSQL & " , " & iMenge3 & " "
    
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    fülleStaffelGruppen Combo3
    
    Combo3.Text = sStaffName & Space(100) & lStaffNr
        
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherStaffelGruppe"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Function ermMaxStaffNr() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    ermMaxStaffNr = 0
    
    sSQL = "Select max(staffnr) as Maxi from STAFFEL_KVK_GRUPPE"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        ermMaxStaffNr = rsrs!maxi
    End If
    rsrs.Close: Set rsrs = Nothing
    
    ermMaxStaffNr = ermMaxStaffNr + 1
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermMaxStaffNr"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub UnsichtbarSG(bSichtbar As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Visible = bSichtbar
    Text1(1).Visible = bSichtbar
    Text1(2).Visible = bSichtbar
    
    Text1(3).Visible = bSichtbar
    Text1(4).Visible = bSichtbar
    Text1(5).Visible = bSichtbar
    Text1(6).Visible = bSichtbar
    
    Label1(4).Visible = bSichtbar
    Label1(5).Visible = bSichtbar
    Label1(6).Visible = bSichtbar
    
    Command2(5).Visible = bSichtbar
    Command2(7).Visible = bSichtbar
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UnsichtbarSG"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Artikel_zur_Gruppe_hinzufuegen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim sStaffNr    As String
    Dim lartnr      As Long
    
    lartnr = Val(Label1(1).Caption)
    
    If lartnr = 0 Then
        Exit Sub
    End If

    sStaffNr = Right(Combo3.Text, 4)
    
    If Val(sStaffNr) = 0 Then
        MsgBox "Bitte wählen Sie eine Staffelgruppe aus!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If

    sSQL = "Delete from Staffel_KVK_Artikel where ARTNR = " & lartnr
'    sSQL = sSQL & " and Staffnr = " & Val(sStaffnr)
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Staffel_KVK_Artikel (Artnr,Staffnr) values  "
    sSQL = sSQL & " ( " & lartnr
    sSQL = sSQL & " , " & Val(sStaffNr)
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    ZeigeArtikel_StaffelGruppe Val(sStaffNr), List2
        
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Artikel_zur_Gruppe_hinzufuegen"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub deleteStaffelpreis(Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim bFound      As Boolean
    Dim sArtnr      As String
    Dim lcount      As Long
    
    bFound = False

    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    
    If Not bFound Then
        MsgBox "Bitte markieren Sie einen Artikel!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    
    
    
    Dim sStaffNr As String
    sStaffNr = Right(Combo3.Text, 4)
    
    If Val(sStaffNr) = 0 Then
        MsgBox "Bitte wählen Sie eine Staffelgruppe aus!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
        
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            sArtnr = Trim(Left(List2.list(lcount), 6))
            
            sSQL = "Delete from STAFFEL_KVK_ARTIKEL where StaffNr = " & Val(sStaffNr) & " And artnr = " & sArtnr
            gdBase.Execute sSQL, dbFailOnError
        End If
    Next lcount
    
    ZeigeArtikel_StaffelGruppe Val(sStaffNr), Listx
    
    
    
    'gruppenzugehörikeit klären
    Gruppenzugehoerigkeit Label1(1).Caption
    
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "deleteStaffelpreis"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub loescheKompletteGruppe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim sStaffNr    As String
    Dim iRet        As Integer
    
    sStaffNr = Right(Combo3.Text, 4)
    
    If Val(sStaffNr) = 0 Then
        MsgBox "Bitte wählen Sie eine Staffelgruppe aus!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    iRet = MsgBox("Möchten Sie wirklich die komplette Staffelgruppe löschen", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
    If iRet = vbNo Then
        Exit Sub
    End If
        
    sSQL = "Delete from STAFFEL_KVK_ARTIKEL where StaffNr = " & Val(sStaffNr)
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from STAFFEL_KVK_GRUPPE where StaffNr = " & Val(sStaffNr)
    gdBase.Execute sSQL, dbFailOnError
        
    List2.Clear
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loescheKompletteGruppe"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    If gsARTNR = "" Then
        Exit Sub
    End If
    
    fülleStaffelGruppen Combo3
    
    
    
    

    Label1(1).Caption = gsARTNR
    
    Gruppenzugehoerigkeit Label1(1).Caption
    
    
    
    
    
    
    UnsichtbarSG False
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil StaffelpreiseKVK ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Gruppenzugehoerigkeit(sArtnr As String)
On Error GoTo LOKAL_ERROR

    Dim sGruppe As String
    Dim sGruppenNr As String
    
    sGruppe = ermGruppe(Label1(1).Caption)
    
    If sGruppe <> "" Then
    
        sGruppenNr = ermGruppenNr(sGruppe)
    
        Label1(7).Caption = "befindet sich in dieser Gruppe: " & sGruppe
        Label1(7).Visible = True
        Label1(7).ForeColor = glWarn
        
        Combo3.Text = sGruppe & Space(100) & sGruppenNr
        Command2(2).Visible = False
        
    Else
        Label1(7).Caption = ""
        Label1(7).Visible = False
        Command2(2).Visible = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Gruppenzugehoerigkeit"
    Fehler.gsFehlertext = "Im Programmteil StaffelpreiseKVK ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermGruppe(sArtnr As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim sStaffNr As String
    
    ermGruppe = ""

    sSQL = "Select * from STAFFEL_KVK_Artikel where artnr = " & sArtnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!STAFFNR) Then
            sStaffNr = rsrs!STAFFNR
        End If
    
    End If
    rsrs.Close
    
    
    If Val(sStaffNr) > 0 Then
        sSQL = "Select * from STAFFEL_KVK_Gruppe where staffnr = " & sStaffNr
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!StaffName) Then
                ermGruppe = rsrs!StaffName
            End If
        
        End If
        rsrs.Close
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermGruppe"
    Fehler.gsFehlertext = "Im Programmteil StaffelpreiseKVK ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermGruppenNr(sGruppenbez As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermGruppenNr = "0"

    sSQL = "Select Staffnr from STAFFEL_KVK_Gruppe where StaffName = '" & sGruppenbez & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!STAFFNR) Then
            ermGruppenNr = rsrs!STAFFNR
        End If
    
    End If
    rsrs.Close
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermGruppenNr"
    Fehler.gsFehlertext = "Im Programmteil StaffelpreiseKVK ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub fülleStaffelGruppen(cbox As ComboBox)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    cbox.Clear
    cbox.AddItem "bitte auswählen"
    cbox.Text = "bitte auswählen"

    sSQL = "Select * from STAFFEL_KVK_GRUPPE order by staffname"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!StaffName) Then
                cbox.AddItem rsrs!StaffName & Space(100) & rsrs!STAFFNR
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fülleStaffelGruppen"
    Fehler.gsFehlertext = "Im Programmteil StaffelpreiseKVK ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command2_Click 0
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    Select Case Index
    
        Case 1, 4, 5
            cValid = "1234567890," & Chr$(8)
            cZeichen = Chr$(KeyAscii)
            If InStr(Text1(Index).Text, ",") > 0 And cZeichen = "," Then
                KeyAscii = 0
            End If
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 0, 3, 6
            cValid = "1234567890" & Chr$(8)
            cZeichen = Chr$(KeyAscii)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Staffelpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub






