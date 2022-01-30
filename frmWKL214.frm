VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWKL214 
   Caption         =   "LEK-Vergleich"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
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
      Height          =   355
      Index           =   0
      Left            =   8160
      MaxLength       =   6
      TabIndex        =   13
      ToolTipText     =   "Sind Ihre Lieferantenkürzel ordentlich gepflegt, so können Sie mit dem Kürzel arbeiten. Geben Sie z.B. Joop ein!"
      Top             =   360
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
      Height          =   355
      Index           =   7
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   11
      ToolTipText     =   "Sind Ihre Lieferantenkürzel ordentlich gepflegt, so können Sie mit dem Kürzel arbeiten. Geben Sie z.B. Joop ein!"
      Top             =   360
      Width           =   1215
   End
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   12
      Left            =   11280
      TabIndex        =   7
      Top             =   360
      Width           =   345
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
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
      Caption         =   "?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.PictureBox picprogress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   9315
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   6855
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   12615
      Begin VB.CheckBox Check1 
         Caption         =   "nur vergleichbare Artikel heranziehen, die bei beiden Lieferanten enthalten sind."
         Height          =   1095
         Left            =   9600
         TabIndex        =   18
         Top             =   840
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   11033
         _Version        =   393216
         ForeColorSel    =   8454143
         FocusRect       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   9
         Top             =   240
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
         Caption         =   "Vergleichen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   21
         Top             =   5280
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
         Caption         =   "Drucken"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   9600
         TabIndex        =   17
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Anzahl der Artikel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   16
         Top             =   5880
         Width           =   2055
      End
   End
   Begin VB.TextBox txtStatus 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command11 
      Height          =   360
      Left            =   10800
      TabIndex        =   8
      Top             =   360
      Width           =   405
      _ExtentX        =   0
      _ExtentY        =   0
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
      ToolTip         =   "Spaltenanordung der Tabelle bestimmen"
      ToolTipTitle    =   "Spaltenanordung"
      ButtonStyle     =   2
      Caption         =   ""
      Filename        =   "D:\Thomas\VB6\Winkiss\Zubehör\tab24.gif"
      Picture         =   "frmWKL214.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   360
      Index           =   6
      Left            =   5880
      TabIndex        =   10
      Top             =   360
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
      Height          =   360
      Index           =   0
      Left            =   9480
      TabIndex        =   12
      Top             =   360
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "2. Lieferant:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6960
      TabIndex        =   20
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "1. Lieferant:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   19
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "kein Lieferant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "LEK-Vergleich"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      Left            =   120
      TabIndex        =   1
      Top             =   7920
      Width           =   9255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "kein Lieferant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   15
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmWKL214"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerArtnr          As Byte

Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
    
        Case 6
            Text1_KeyUp 7, vbKeyF2, 0
        Case 0
            Text1_KeyUp 0, vbKeyF2, 0
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Artnr"
    gsZSpalte1 = "Farbnr"
    gstab = "LEKVERGLEICH"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cFarbkenn   As String
    Dim iRet        As Integer
    Dim ctmp        As String
    Dim lcount      As Long
    Dim i           As Integer

    Select Case Index
        Case 0
            Unload frmWKL214
        Case 1
            'suchen
            
            ermVergleich_Artikel Text1(7).Text, Text1(0).Text
            
            ZeigeArtikel192
            
        Case 2 'Drucken
            Drucke_LEK_VERGLEICH
            
        Case 12
            gsHelpstring = "LEK-Vergleich"
            frmWKL110.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub ZeigeArtikel192()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    
    MSFlexGrid1.Clear
    
    Command5(2).Enabled = False
    
    If Not NewTableSuchenDBKombi("art214", gdBase) Then
        anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
        Exit Sub
    Else
        If Datendrin("art214", gdBase) = False Then
            anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
            Exit Sub
        End If
    End If
    
    Command5(2).Enabled = True
    
    anzeige "normal", "Artikel werden angezeigt, bitte warten...", Label1(4)
    
    Screen.MousePointer = 11
    
    Tabcheck "LEKVERGLEICH"
    FormatGridOverTablay "LEKVERGLEICH"

    With MSFlexGrid1
        .Redraw = False
'        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
            aBreite(j) = TextWidth(.TextMatrix(0, j)) '* 1.8
        Next j
    End With
    
    ermittlespalten
    
    GridFuellen "Select * from art214 order by artnr "
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    

    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeArtikel192"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Drucke_LEK_VERGLEICH()
    On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim i As Integer

    Screen.MousePointer = 11

    loeschNEW "ART214PRINT", gdBase
    CreateTableT2 "ART214PRINT", gdBase

    cSQL = "Insert into ART214PRINT select * from ART214  "
    gdBase.Execute cSQL, dbFailOnError

    anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)

    reportbildschirm "", "aWKL214" '"aZEN192a"

    anzeige "normal", "", Label1(4)

    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke_LEK_VERGLEICH"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub GridFuellen(cSQL As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim iRet        As Integer
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim lMax        As Long
    Dim lAnz        As Long
    
    If cSQL = "" Then
        Exit Sub
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    picprogress.Visible = True
    With MSFlexGrid1
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMax = rsrs.RecordCount
        lAnz = lMax
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            .Rows = lrow + 1
            .Col = 0
            
            txtStatus.Text = (lrow * 100) / lMax
            
            Select Case lMax
                Case Is > 5000
                
                    j = lAnz Mod 500
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label1(4)
                    End If
                
                Case Is > 1000
                
                    j = lAnz Mod 100
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label1(4)
                    End If
                
                Case Is <= 500
                
                    j = lAnz Mod 50
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label1(4)
                    End If
        
            End Select
    
            lAnz = lAnz - 1
            
            For i = 0 To byAnzahlSpalten - 1
                .Row = 0
                .Col = i
                
                If sSpaltenname(i) = .Text Then
                    
                    Select Case UCase(sSpaltenname(i))
                        Case Is = "LEK 1", "LEK 2", "DIFFERENZ", "KVK"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                            
                        Case Is = "HS 1", "HS 2"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = Fix(rsrs(sSpaltenbez(i)))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = sWert
                            
                            If CDbl(sWert) >= 100 Then
                                .CellBackColor = &H8000&
                            ElseIf CDbl(sWert) > 79.99 Then
                                .CellBackColor = &HC000&
                            ElseIf CDbl(sWert) > 59.99 Then
                                .CellBackColor = &HFF00&
                            ElseIf CDbl(sWert) > 39.99 Then
                                .CellBackColor = &HFFFF&
                            ElseIf CDbl(sWert) > 19.99 Then
                                .CellBackColor = &H80C0FF
                            ElseIf CDbl(sWert) > 0 Then
                                .CellBackColor = &H80FF&
                            ElseIf CDbl(sWert) <= 0 Then
                                .CellBackColor = &HFF&
                            End If
                        
                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = ""
                            End If
                            .Row = lrow
                            .Text = sWert
                            
                    End Select
                    
                    If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
                        aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                    End If
                    
                End If
            Next i
                                
            rsrs.MoveNext
        Loop
        
        Frame2.Visible = True
        
        anzeige "normal", CStr(lMax), Label1(3)
        anzeige "normal", lMax & " Artikel", Label1(4)
    
    Else
        Frame2.Visible = False
        anzeige "normal", "", Label1(3)
        anzeige "rot", "Es wurden keine Artikel ermittelt.", Label1(4)
    End If
    
    For i = 0 To byAnzahlSpalten - 1
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.8
    Next i
    
        
    rsrs.Close
    If byAnzahlSpalten < 2 Then
    Else
        .FixedCols = 1
    End If
    
    picprogress.Visible = False
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    
    
    .RowHeight(1) = 0
    lrow = lrow - 1
    .Redraw = True
'    .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."

    Fehlermeldung1
  
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim cVon As String
    Dim cBis As String
    
    Screen.MousePointer = 11
    
    PositionierenWKL192
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    Me.Refresh
    
   
    
    Frame2.Visible = True
    
    anzeige "normal", "", Label1(4)
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
    Dim i           As Long
    Dim j           As Long
    
    With gridx
    
        ReDim bBreit(.Cols - 1)
        
        For j = 0 To .Rows - 1
            For i = 0 To .Cols - 1
                If TextWidth(.TextMatrix(j, i)) > bBreit(i) Then
                    bBreit(i) = TextWidth(.TextMatrix(j, i))
                End If
            Next i
        Next j
        
        Select Case Screen.Height
            Case Is > 15000
                siFak = 1.5
            Case Is > 12000
                siFak = 1.4
            Case Is > 11000
                siFak = 1.2
            Case Is > 10000
                siFak = 1.1
            Case Is > 8000
                siFak = 1#
        End Select
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = bBreit(i) * siFak * siEigFak
        Next i
    
    End With
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Farbanpassung(cFabm As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
    Screen.MousePointer = 11
    
    sSQL = "update art45 set farbnr = " & Val(cFabm) & " "
    gdBase.Execute sSQL, dbFailOnError
    
    BringFarbeInsSpiel "Art45", gdBase
    
    sSQL = "update artikel inner join art45 on artikel.artnr = art45.artnr"
    sSQL = sSQL & " set AWM = '" & cFabm & "'"
    sSQL = sSQL & " , LASTDATE = '" & DateValue(Now) & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Farbanpassung"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL192()
On Error GoTo LOKAL_ERROR

    With Frame2
        .Top = 960
        .Height = 6735
        .Width = 11775
        .Left = 0

    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL192"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    loeschNEW "ART214", gdBase
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
Private Sub ermVergleich_Artikel(sLinr1 As String, sLinr2 As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim siAnzeige As Single
    
    
    If sLinr1 = sLinr2 Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "ART214", gdBase
    CreateTableT2 "ART214", gdBase
    
    sSQL = "Create index artnr on ART214(artnr) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "die Artikel werden ermittelt...", Label1(4)
    
    sSQL = " Insert into ART214 select ARTNR from Artlief where linr = " & sLinr1
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 12
    
    loeschNEW "ART214_DAT2", gdBase
    
    sSQL = "select ARTNR, lekpr into ART214_DAT2 "
    sSQL = sSQL & " from Artlief where linr = " & sLinr2
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 14
    
    sSQL = " Alter table ART214_DAT2 add erkannt varchar(1) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 15
    
    sSQL = " Update ART214_DAT2 set erkannt = 'N'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 16
    
    sSQL = "Create index artnr on ART214_DAT2(artnr) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 18
    
    sSQL = " Update ART214_DAT2 inner join ART214 on ART214_DAT2.artnr = ART214.artnr "
    sSQL = sSQL & " set erkannt = 'J'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 20

    sSQL = " Insert into ART214 select ARTNR"
    sSQL = sSQL & " from ART214_DAT2"
    sSQL = sSQL & " where ART214_DAT2.erkannt = 'N'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = " Update ART214 set LEK_1 = 0"
    sSQL = sSQL & ", LEK_2 = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    sSQL = " Update ART214 inner join Artlief on Art214.artnr = Artlief.artnr "
    sSQL = sSQL & " set LEK_1 = Artlief.lekpr "
    sSQL = sSQL & " , RKZ_1 = Artlief.rkz where linr = " & sLinr1
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    sSQL = " Update ART214 inner join Artlief on Art214.artnr = Artlief.artnr "
    sSQL = sSQL & " set LEK_2 = Artlief.lekpr "
    sSQL = sSQL & " , RKZ_2 = Artlief.rkz where linr = " & sLinr2
    gdBase.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 55
    
    sSQL = " Update ART214 set LASTVK = null"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "Last_VK" & srechnertab, gdBase
    
    txtStatus.Text = 62
    
    sSQL = "Select Kassjour.Artnr, Max(adate) as LASTVK into Last_VK" & srechnertab
    sSQL = sSQL & " from KASSJOUR inner join ART214 on KASSJOUR.ARTNR = ART214.Artnr  "
    sSQL = sSQL & " group by  KASSJOUR.ARTNR "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 71
    
    sSQL = "Update ART214 inner join Last_VK" & srechnertab & " on ART214.Artnr = Last_VK" & srechnertab & ".Artnr "
    sSQL = sSQL & " Set ART214.LASTVK = Last_VK" & srechnertab & ".LASTVK  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "Last_VK" & srechnertab, gdBase
    
    
    
    
    
    
    txtStatus.Text = 75
    
    
            
            
    
    sSQL = " Update ART214 inner join Artikel on Art214.artnr = Artikel.artnr "
    sSQL = sSQL & " set ART214.Bezeich = Artikel.Bezeich "
    sSQL = sSQL & " , ART214.KVKPR1 = Artikel.KVKPR1 "
    sSQL = sSQL & " , ART214.BESTAND = Artikel.BESTAND "
    sSQL = sSQL & " , ART214.MWST = Artikel.MWST "
    sSQL = sSQL & " , ART214.EAN = Artikel.EAN "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ART214 set HS_1 = "
    sSQL = sSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStV & ")) - LEK_1) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStV & ")) "
    sSQL = sSQL & "  "
    sSQL = sSQL & "  where MWST = 'V' and LEK_1 > 0 and KVKPR1 > 0"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ART214 set HS_2 = "
    sSQL = sSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStV & ")) - LEK_2) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStV & ")) "
    sSQL = sSQL & "  "
    sSQL = sSQL & "  where MWST = 'V' and LEK_2 > 0 and KVKPR1 > 0"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    sSQL = "Update ART214 set HS_1 = "
    sSQL = sSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStE & ")) - LEK_1) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStE & ")) "
    sSQL = sSQL & "  "
    sSQL = sSQL & "  where MWST = 'E' and LEK_1 > 0 and KVKPR1 > 0"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ART214 set HS_2 = "
    sSQL = sSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStE & ")) - LEK_2) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStE & ")) "
    sSQL = sSQL & "  "
    sSQL = sSQL & "  where MWST = 'E' and LEK_2 > 0 and KVKPR1 > 0"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
            
            
            
    sSQL = "Update ART214 set HS_1 = "
    sSQL = sSQL & "  ((((KVKPR1 * 100) / (100)) - LEK_1) *100) / ((KVKPR1 * 100) / (100)) "
    sSQL = sSQL & "  "
    sSQL = sSQL & "  where MWST = 'O' and LEK_1 > 0 and KVKPR1 > 0"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ART214 set HS_2 = "
    sSQL = sSQL & "  ((((KVKPR1 * 100) / (100)) - LEK_2) *100) / ((KVKPR1 * 100) / (100)) "
    sSQL = sSQL & "  "
    sSQL = sSQL & "  where MWST = 'O' and LEK_2 > 0 and KVKPR1 > 0"
    gdBase.Execute sSQL, dbFailOnError

    
    
    
    If Check1.Value = vbChecked Then
    
        txtStatus.Text = 76
    
        sSQL = "Delete from  ART214 where lek_1 = 0"
        gdBase.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 78
    
        sSQL = "Delete from  ART214 where lek_2= 0"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    
    
    
    
    txtStatus.Text = 95
   
    sSQL = " Update ART214 set Differenz = LEK_1 - LEK_2 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
 
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermVergleich_Artikel"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub



Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    
    If KeyCode = vbKeyF2 Then
        lrow = MSFlexGrid1.Row
        gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
        If gsARTNR <> "" Then
            
            frmWKL10.Show 1
            Me.Refresh
            Screen.MousePointer = 11

            MSFlexGrid1.TopRow = lrow
            MSFlexGrid1.Col = SpaltennummerArtnr
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
            
            Screen.MousePointer = 0
        End If
        gsARTNR = ""
    
    End If
    
    MSFlexGrid1.Redraw = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row = 1 Then
        sortierenGrid MSFlexGrid1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

'Private Sub MSFlexGrid1_SelChange()
'On Error GoTo LOKAL_ERROR
'
'Dim cART As String
'
'cART = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
'
'If cART <> "" Then
'    If IsNumeric(cART) Then
'
'    Label1(9).Caption = ErmlzVK(cART)
'    Label1(11).Caption = ErmlzZugang(cART)
'    Label1(13).Caption = ErmFirstZugang(cART)
'
'    End If
'End If
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "MSFlexGrid1_SelChange"
'    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If Index = 0 Then
        LiefKuerzelAufloesung Label1(1), Text1(0)
    End If

    If Index = 7 Then
        LiefKuerzelAufloesung Label1(0), Text1(7)
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 0
                gF2Prompt.cFeld = "LINR"
                
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    Label1(1).Caption = gF2Prompt.cWert
                End If
                
            Case Is = 7
                gF2Prompt.cFeld = "LINR"
                
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    Label1(0).Caption = gF2Prompt.cWert
                End If
        End Select
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil LEK-Vergleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtStatus_Change()
On Error GoTo LOKAL_ERROR

    Dim nProz As Long
  
    nProz = Val(txtStatus.Text)
    ShowProgress picprogress, nProz, 0, 100, True
    picprogress.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtStatus_Change"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub




