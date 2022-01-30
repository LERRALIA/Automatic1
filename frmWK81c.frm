VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK81c 
   BackColor       =   &H00C0C000&
   Caption         =   "Termine - Stornotext"
   ClientHeight    =   8910
   ClientLeft      =   1935
   ClientTop       =   2475
   ClientWidth     =   11910
   Icon            =   "frmWK81c.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8910
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List1 
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
      ItemData        =   "frmWK81c.frx":0442
      Left            =   240
      List            =   "frmWK81c.frx":0444
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   13
      Top             =   2400
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   240
      MaxLength       =   32
      TabIndex        =   12
      Top             =   960
      Width           =   5175
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   6
      Left            =   2640
      TabIndex        =   11
      Top             =   1800
      Width           =   735
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
      Caption         =   "V"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   10
      Top             =   5040
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
      Caption         =   "Leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.CheckBox Check29 
      Caption         =   "Termine auf Bon ohne ""Warnung"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      TabIndex        =   9
      Top             =   5640
      Width           =   5175
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   9720
      TabIndex        =   8
      Top             =   5040
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
      Caption         =   "Leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   8880
      TabIndex        =   5
      Top             =   1800
      Width           =   735
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
      Caption         =   "V"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   6600
      MaxLength       =   32
      TabIndex        =   4
      Top             =   960
      Width           =   5175
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   9960
      TabIndex        =   3
      Top             =   6840
      Width           =   1815
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
      Caption         =   "Test"
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
      ItemData        =   "frmWK81c.frx":0446
      Left            =   6600
      List            =   "frmWK81c.frx":0448
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   2
      Top             =   2400
      Width           =   5175
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9960
      TabIndex        =   1
      Top             =   7440
      Width           =   1815
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9960
      TabIndex        =   0
      Top             =   8040
      Width           =   1815
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "1. Teil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6600
      TabIndex        =   17
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "1. Teil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Stornotext, pro Zeile 32 Zeichen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "noch 32 Zeichen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "noch 32 Zeichen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Stornotext, pro Zeile 32 Zeichen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6600
      TabIndex        =   6
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "frmWK81c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            speichern_Storno_TextT1
            speichern_Storno_Text
            speicher2Druck
        Case 1      'Beenden
            Unload frmWK81c
        Case 2
            DruckeZweitBonAusListe List1, True 'dann drucken
            DruckeZweitBonAusListe Listx, True 'dann drucken
        Case 3
            hinzu
        Case 4
            leeren
        Case 5
            leeren1
        Case 6
            hinzu1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicher2Druck()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    If Check29.Value = vbChecked Then
        sSQL = "Update WKEINSTE Set TNW = true "
        gdApp.Execute sSQL, dbFailOnError
        gbTerminNoWarn = True
    Else
        sSQL = "Update WKEINSTE Set TNW = false "
        gdApp.Execute sSQL, dbFailOnError
        gbTerminNoWarn = False
    End If



    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher2Druck"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub hinzu()
    On Error GoTo LOKAL_ERROR
    
    If Text1(0).Text = "" Then
        Exit Sub
    End If
    
    Listx.AddItem Text1(0).Text
    Text1(0).Text = ""
    Listx.Refresh
    Text1(0).SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "hinzu"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub leeren()
    On Error GoTo LOKAL_ERROR
    
    Dim bFound As Boolean
    Dim lcount As Long
    Dim lAnz As Long
    
    For lcount = 0 To Listx.ListCount - 1
        If Listx.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount
            
    If bFound Then
    
        lAnz = Listx.ListCount - 1
        
        For lcount = 0 To lAnz
    
            If lcount > lAnz Then
                Exit For
            End If
            
            If Listx.Selected(lcount) = True Then
                Listx.RemoveItem lcount
                
                lcount = lcount - 1
                lAnz = lAnz - 1
            End If
        Next lcount
    Else
    
        Listx.Clear
    End If
    
    Listx.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leeren"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub hinzu1()
    On Error GoTo LOKAL_ERROR
    
    If Text1(1).Text = "" Then
        Exit Sub
    End If
    
    List1.AddItem Text1(1).Text
    Text1(1).Text = ""
    List1.Refresh
    Text1(1).SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "hinzu1"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub leeren1()
    On Error GoTo LOKAL_ERROR
    
    Dim bFound As Boolean
    Dim lcount As Long
    Dim lAnz As Long
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount
            
    If bFound Then
    
        lAnz = List1.ListCount - 1
        
        For lcount = 0 To lAnz
    
            If lcount > lAnz Then
                Exit For
            End If
            
            If List1.Selected(lcount) = True Then
                List1.RemoveItem lcount
                
                lcount = lcount - 1
                lAnz = lAnz - 1
            End If
        Next lcount
    Else
    
        List1.Clear
    End If
    
    List1.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leeren1"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lade_Standard_Storno_Text_Teil2()
    On Error GoTo LOKAL_ERROR
    
    Dim cDaten As String
    Dim lcount As Long
    
    Listx.Clear
    
    If NewTableSuchenDBKombi("STORNOTEXT", gdBase) = False Then
    
        cDaten = "Ansonsten müssen wir Ihnen"
        Listx.AddItem cDaten
            
        cDaten = "70 % des Behandlungspreises ver-"
        Listx.AddItem cDaten
            
        cDaten = "rechnen. Wir danken für Ihr Ver-"
        Listx.AddItem cDaten
            
        cDaten = "ständnis!"
        Listx.AddItem cDaten
    Else
        lese_Storno_Text_in_Array
        
        For lcount = LBound(sStornoText) To UBound(sStornoText)
            Listx.AddItem sStornoText(lcount)
        Next lcount
    End If
        
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lade_Standard_Storno_Text_Teil2"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lade_Standard_Storno_Text_Teil1()
    On Error GoTo LOKAL_ERROR
    
    Dim cDaten As String
    Dim lcount As Long
    
    List1.Clear
    
    If NewTableSuchenDBKombi("STORNOTEXTT1", gdBase) = False Then
    
        cDaten = "Bitte kommen Sie rechtzeitig vor"
        List1.AddItem cDaten
        
        cDaten = "Ihrem Behandlungstermin. Ver-"
        List1.AddItem cDaten
    
        cDaten = "spätungen haben leider eine"
        List1.AddItem cDaten
    
        cDaten = "kürzere Behandlung zur Folge."
        List1.AddItem cDaten
    
        cDaten = "Wenn Sie einen Termin nicht"
        List1.AddItem cDaten
        
        cDaten = "einhalten können, bitten wir Sie"
        List1.AddItem cDaten
        
        cDaten = "mindestens 3 Tage davor "
        List1.AddItem cDaten
        
        cDaten = "abzusagen."
        List1.AddItem cDaten

    Else
        lese_Storno_Text_in_Array_T1
        
        For lcount = LBound(sStornoTextT1) To UBound(sStornoTextT1)
            List1.AddItem sStornoTextT1(lcount)
        Next lcount
    End If
        
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lade_Standard_Storno_Text_Teil1"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speichern_Storno_Text()
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "STORNOTEXT", gdBase
    CreateTableT2 "STORNOTEXT", gdBase
    
    Dim lcount As Long
    Dim sSQL As String
    
    For lcount = 0 To Listx.ListCount - 1
        sSQL = "Insert into STORNOTEXT (ZNR,ZTEXT) values (" & lcount & " , '" & Listx.list(lcount) & "')"
        gdBase.Execute sSQL, dbFailOnError
    Next lcount
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichern_Storno_Text"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speichern_Storno_TextT1()
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "STORNOTEXTT1", gdBase
    CreateTableT2 "STORNOTEXTT1", gdBase
    
    Dim lcount As Long
    Dim sSQL As String
    
    For lcount = 0 To List1.ListCount - 1
        sSQL = "Insert into STORNOTEXTT1 (ZNR,ZTEXT) values (" & lcount & " , '" & List1.list(lcount) & "')"
        gdBase.Execute sSQL, dbFailOnError
    Next lcount
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichern_Storno_TextT1"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    lade_Standard_Storno_Text_Teil2
    lade_Standard_Storno_Text_Teil1
    
    If gbTerminNoWarn Then
        Check29.Value = vbChecked
    Else
        Check29.Value = vbUnchecked
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_Change(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case Index
    
        Case 0
            anzeige "normal", "noch " & 32 - Len(Text1(0).Text) & " Zeichen", Label1(0)
        Case 1
            anzeige "normal", "noch " & 32 - Len(Text1(1).Text) & " Zeichen", Label1(1)
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command1_Click 3
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub


