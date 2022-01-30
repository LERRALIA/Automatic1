VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWKL193 
   Caption         =   "gelöschte Termine"
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
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   9315
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   6735
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   10200
         MaxLength       =   8
         TabIndex        =   20
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   1695
         Left            =   9600
         TabIndex        =   9
         Top             =   2040
         Width           =   1935
         Begin VB.OptionButton Option1 
            Caption         =   "aktuelles Jahr"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "letzten 7 Tage"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Gestern"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Heute"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   11668
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
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   0
         Left            =   11160
         TabIndex        =   15
         ToolTipText     =   "Kalender"
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
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
         ToolTip         =   "Wählen Sie hier das Datum aus."
         ToolTipTitle    =   "Kalender"
         ButtonStyle     =   2
         Caption         =   ""
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   1
         Left            =   11160
         TabIndex        =   16
         ToolTipText     =   "Kalender"
         Top             =   720
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
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
         ToolTip         =   "Wählen Sie hier das Datum aus."
         ToolTipTitle    =   "Kalender"
         ButtonStyle     =   2
         Caption         =   ""
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   17
         Top             =   3840
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   23
         Top             =   4440
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "KundNr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   10200
         TabIndex        =   21
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   9720
         TabIndex        =   18
         Top             =   720
         Width           =   375
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
      Left            =   9600
      TabIndex        =   3
      Top             =   360
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
      Picture         =   "frmWKL192.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "gelöschte Termine"
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
      TabIndex        =   2
      Top             =   120
      Width           =   8775
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
End
Attribute VB_Name = "frmWKL193"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerBEGINDAT       As Byte
Dim SpaltennummerADATE          As Byte
Dim SpaltennummerAZEIT          As Byte


Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 0
            Text1(0).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            Text1(1).SetFocus
        Case Is = 1
            Text1(1).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            'fertig
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Bediener"
    gstab = "TERMDEL"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL193
        Case 1
            erm_Term Text1(0).Text, Text1(1).Text, Text1(2).Text
            ZeigeTERMDEL193
        Case 2 'löschen
        
            If MSFlexGrid1.RowSel > 1 Then
            
            
            
            
                FlexGrid_Delete MSFlexGrid1
            
            
                

            Else
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
            End If
        
            erm_Term Text1(0).Text, Text1(1).Text, Text1(2).Text
            ZeigeTERMDEL193
        Case 12
            gsHelpstring = "gelöschte Termine"
            frmWKL110.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FlexGrid_Delete(oGrid As MSFlexGrid)
On Error GoTo LOKAL_ERROR

    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
    Dim lBig As Long
    
    Dim cArtNr As String
  

    With oGrid
        ' aktuelle Selektion merken
      
        nRow = .Row
        nCol = .Col
        nRowSel = .RowSel
        nColSel = .ColSel
      
      
        If nRow > nRowSel Then
            lBig = nRow
            nDelRow = nRowSel - 1
        Else
            lBig = nRowSel
            nDelRow = nRow - 1
        End If
      
      
        Do While nDelRow < lBig
        
            nDelRow = nDelRow + 1
            
            If nDelRow > 1 Then
            
            
            
            
                
            
                Dim sZeit As String
                Dim sDatum As String
                Dim lDat As Long
                Dim cdat As String


                sDatum = Trim(MSFlexGrid1.TextMatrix(nDelRow, SpaltennummerADATE))
                sZeit = Trim(MSFlexGrid1.TextMatrix(nDelRow, SpaltennummerAZEIT))

                lDat = DateValue(sDatum)
                cdat = Trim$(Str$(lDat))


                Dim sSQL As String
                sSQL = "Delete from TERMDEL where azeit = '" & sZeit & "'"
                sSQL = sSQL & " and adate = " & cdat & ""
                gdBase.Execute sSQL, dbFailOnError
            
            
            
'                cArtNr = Trim(.TextMatrix(nDelRow, SpaltennummerArtnr))
'                LoescheArtikelWKL10 cArtNr
                .RowHeight(nDelRow) = 0
                
            End If
        Loop

  
  End With
  

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Delete"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeTERMDEL193()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    
    MSFlexGrid1.Clear
    
    If Not NewTableSuchenDBKombi("TERMDEL193", gdBase) Then
        anzeige "rot", "Es sind keine gelöschte Termine ermittelt worden.", Label1(4)
        Exit Sub
    Else
        If Datendrin("TERMDEL193", gdBase) = False Then
            anzeige "rot", "Es sind keine gelöschte Termine ermittelt worden.", Label1(4)
            Exit Sub
        End If
    End If
    
    anzeige "normal", "gelöschte Termine werden angezeigt, bitte warten...", Label1(4)
    
    Screen.MousePointer = 11
    
    Tabcheck "TERMDEL"
    FormatGridOverTablay "TERMDEL"

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
    
    
    GridFuellen "Select * from TERMDEL193 order by BEGINDAT desc,BEGINZEIT desc"
    
    
'    GridFuellen "Select * from TERMDEL193 order by adate desc,AZEIT desc"
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.1 * gdTabfak
    

    MSFlexGrid1.Redraw = True
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeArtikel192"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."
    
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
    .Visible = False

    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMax = rsrs.RecordCount
        lAnz = lMax


'        Anzeige "normal", "Es werden " & lMax & " Artikel angezeigt...", Label1(4)
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
                        Case Is = "LEK", "KVK", "LUG", "LEK-WERT", "KVK-WERT"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")

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

        
        anzeige "normal", lMax & " gelöschte Termine", Label1(4)

    Else
        Frame2.Visible = False
        
        anzeige "rot", "Es wurden keine gelöschte Termine ermittelt.", Label1(4)


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
    .Visible = True
    End With

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    PositionierenWKL193
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    Me.Refresh
    
    Option1_Click 5
    
    erm_Term Text1(0).Text, Text1(1).Text, Text1(2).Text
    ZeigeTERMDEL193
    
    Frame2.Visible = True
    
    anzeige "normal", "", Label1(4)
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "BEGINDAT"
                SpaltennummerBEGINDAT = i
            Case Is = "ADATE"
                SpaltennummerADATE = i
            Case Is = "AZEIT"
                SpaltennummerAZEIT = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
    Dim i           As Integer
    Dim j           As Integer
    
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
Private Sub PositionierenWKL193()
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
    Fehler.gsFunktion = "PositionierenWKL193"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    loeschNEW "ART192", gdBase
    loeschNEW "ART192PRINT", gdBase
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
Private Sub erm_Term(cVon As String, cBis As String, cKundnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim siAnzeige As Single
    Dim bAnd As Boolean
    bAnd = False
    
    If cKundnr <> "" Then
        If IsNumeric(cKundnr) = False Then
            cKundnr = ""
        End If
    End If
    
    Dim lVon As Long
    Dim lBis As Long
    
    If cVon <> "" And cBis <> "" Then
        lVon = DateValue(cVon)
        lBis = DateValue(cBis)
        
        cVon = Trim$(Str$(lVon))
        cBis = Trim$(Str$(lBis))
    End If
    
    Screen.MousePointer = 11
    
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "TERMDEL193", gdBase
    CreateTableT2 "TERMDEL193", gdBase
    
    anzeige "normal", "die gelöschten Termine werden ermittelt...", Label1(4)
    
    sSQL = " Insert into TERMDEL193 select "
    sSQL = sSQL & " ADATE "
    sSQL = sSQL & ", AZEIT "
    sSQL = sSQL & ", BED "
    sSQL = sSQL & ", KUNDNR  "
    sSQL = sSQL & ", GRUND "
    sSQL = sSQL & ", ERSTBED "
    sSQL = sSQL & ", DAUER "
    sSQL = sSQL & ", BEGINDAT "
    sSQL = sSQL & ", BEGINZEIT "
    sSQL = sSQL & " from TERMDEL "
    
    If cVon <> "" And cBis <> "" Then
        If bAnd = True Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " Where "
        End If
        sSQL = sSQL & "  adate between  " & cVon & " And " & cBis
        bAnd = True
    End If
    
    If cKundnr <> "" Then
    
        If bAnd = True Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " Where "
        End If
        
        sSQL = sSQL & "  Kundnr = " & cKundnr
        bAnd = True
    End If
    
    
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    sSQL = "Update TERMDEL193 inner join Kunden  "
    sSQL = sSQL & " on TERMDEL193.KUNDNR = Kunden.Kundnr "
    sSQL = sSQL & " Set TERMDEL193.NAME = Kunden.NAME "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 56
    
    sSQL = "Update TERMDEL193 inner join Bedname  "
    sSQL = sSQL & " on TERMDEL193.BED = Bedname.bednu "
    sSQL = sSQL & " Set TERMDEL193.BEDNAME = Bedname.BEDNAME "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update TERMDEL193 inner join Bedname  "
    sSQL = sSQL & " on TERMDEL193.ERSTBED = Bedname.bednu "
    sSQL = sSQL & " Set TERMDEL193.ERSTBEDNAME = Bedname.BEDNAME "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 63
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
 
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "erm_Term"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row = 1 Then
    
        Screen.MousePointer = 11
        If MSFlexGrid1.Col = SpaltennummerADATE Then
        
            If byteSortReihen = 1 Then
                byteSortReihen = 2
                GridFuellen "Select * from TERMDEL193 order by adate desc"
            ElseIf byteSortReihen = 2 Then
                byteSortReihen = 1
                GridFuellen "Select * from TERMDEL193 order by adate asc"
            End If
            
        ElseIf MSFlexGrid1.Col = SpaltennummerBEGINDAT Then
        
            If byteSortReihen = 1 Then
                byteSortReihen = 2
                GridFuellen "Select * from TERMDEL193 order by BEGINDAT desc"
            ElseIf byteSortReihen = 2 Then
                byteSortReihen = 1
                GridFuellen "Select * from TERMDEL193 order by BEGINDAT asc"
            End If
        Else
            sortierenGrid MSFlexGrid1
        End If
        Screen.MousePointer = 0
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index

        Case Is = 0 'aktuelles Jahr
        
            Text1(0).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
    
        Case Is = 5     'nächsten 7 Tage
            Text1(0).Text = Format(DateValue(Now) - 7, "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
        Case Is = 6     'gestern
            Text1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now) - 1, "DD.MM.YY")
        Case Is = 7     'heute
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."

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



