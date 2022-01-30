VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmWKL75 
   Caption         =   "EAN Duplikate"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL75.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "doppelte Artnr raus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   6000
         TabIndex        =   28
         Top             =   5280
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "alle Alt-Artikel(weiße Zeilen) löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   8160
         TabIndex        =   25
         Top             =   5280
         Width           =   3495
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Löschen ohne Nachzufragen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   8520
         TabIndex        =   24
         Top             =   120
         Width           =   3135
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   0
         Left            =   9360
         TabIndex        =   19
         Top             =   6480
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
         Caption         =   "EAN"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   1
         Left            =   10200
         TabIndex        =   11
         Top             =   6480
         Width           =   1335
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   2
         Left            =   8160
         TabIndex        =   10
         Top             =   6480
         Width           =   1095
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
         Caption         =   "Details"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   4
         Left            =   8160
         TabIndex        =   1
         Top             =   6960
         Width           =   3375
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
         Caption         =   "Fertig / Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7858
         _Version        =   393216
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   5640
         Width           =   7815
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6840
            MaxLength       =   7
            TabIndex        =   18
            Top             =   1050
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Kassenverkaufspreis dem verbleibenden Artikel zuordnen"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   6495
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   3
            Left            =   6600
            TabIndex        =   8
            Top             =   1280
            Width           =   1095
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
            Caption         =   "Löschen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.CheckBox Check2 
            Caption         =   "die Lieferantendaten übernimmt der verbleibende Artikel"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   7575
         End
         Begin VB.CheckBox Check2 
            Caption         =   "die Verkaufszahlen übernimmt der verbleibende Artikel"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   7335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "den Bestand übernimmt der verbleibende Artikel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   5655
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5880
            MaxLength       =   4
            TabIndex        =   4
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label1 
            Appearance      =   0  '2D
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fest Einfach
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   7
            Left            =   120
            TabIndex        =   23
            Tag             =   "Shape"
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FFFF&
            Caption         =   "Achtung neuen Artikeln kann kein Bestand und Kassenpreis zugewiesen werden"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   7215
         End
         Begin VB.Label Label3 
            Caption         =   "Lösch - und Übernahmeoptionen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   4215
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Notizen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Width           =   11535
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "neue Artikel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   21
         Top             =   5280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   20
         Tag             =   "Shape"
         Top             =   5280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Tag             =   "Shape"
         Top             =   5280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Artikel der Stammdaten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   16
         Top             =   5280
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "doppelte Artikel (EAN-Code) mit Doppelklick unwiderruflich löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   10095
      End
      Begin VB.Label Label1 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   8160
         TabIndex        =   12
         Top             =   5640
         Width           =   3375
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   300
      Index           =   5
      Left            =   10080
      TabIndex        =   26
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
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
      Caption         =   "alle Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
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
      Caption         =   "alle Löschen, die keine Verkäufe aufweisen, keinen Bestand haben und deren Anlagedatum älter als 30 Tage ist"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   5160
      TabIndex        =   27
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "EAN Duplikate"
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
      TabIndex        =   14
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmWKL75"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 5 Then
    If Check2(5).Value = vbChecked Then
    
        iRet = MsgBox("Möchten Sie alle Alt - Artikel (weiße Zeilen) löschen?", vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            
            
        If iRet = vbYes Then
            alleweißenlöschen
        End If
    
    End If
ElseIf Index = 6 Then
    
    'doppelte raus
    'eankoml
    
    
    
    
    'Duplikate löschen
    
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lcount          As Long
    Dim cArtNr          As String
    Dim sSQL            As String
    
    loeschNEW "alit" & srechnertab, gdBase
    sSQL = "select count(Artnr) as count ,Artnr into alit" & srechnertab & " from eankoml group by Artnr having count(Artnr) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "artdupli", gdBase
    sSQL = "Select * into artDupli from eankoml where artnr = -1 "
    gdBase.Execute sSQL, dbFailOnError
    
'    anzeige "normal", "Ermittlung der Duplikate...", Label1(4)
    
    Set rsartDupli = gdBase.OpenRecordset("artDupli", dbOpenTable)
    
    Set rsrs = gdBase.OpenRecordset("alit" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = Trim(rsrs!artnr)
            End If
            
            sSQL = "Select * from eankoml where artnr = " & cArtNr
            Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst
                
                rsArt.MoveNext
                Do While Not rsArt.EOF
                    
                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).Value = rsArt(i).Value
                    Next i
                    rsartDupli.Update
                    
                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close
    
    loeschNEW "alit" & srechnertab, gdBase
    
    
    anzeigenDupliEan_mitVK
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Im Programmteil EAN DUPLIKATE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub alleweißenlöschen()
On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim j               As Integer
    Dim lartnr          As Long
    Dim lrow            As Long
    Dim sEAN            As String
    Dim lLinr           As Long
   
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Redraw = False
    
    Screen.MousePointer = 11
    
    For i = 2 To MSFlexGrid1.Rows - 1
        lrow = i
        MSFlexGrid1.Row = lrow
        
        anzeige "normal", lrow & " Artikel wurden gelöscht...", Label1(1)

        lartnr = MSFlexGrid1.TextMatrix(lrow, 0)
            
        MSFlexGrid1.Row = lrow
        MSFlexGrid1.Col = 2
        
        If MSFlexGrid1.CellBackColor = vbWhite Then
        
            sEAN = MSFlexGrid1.TextMatrix(lrow, 3)
            lLinr = Val(MSFlexGrid1.TextMatrix(lrow, 4))
                
            sSQL = "Select * from eankoml where artnr =" & lartnr
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                rsrs.MoveLast
            
                If rsrs.RecordCount > 1 Then
                
                    'kommt eine ArtNr mehrmals vor, so wird nur die EAN gelöscht
                    
                    sSQL = "Update artikel set EAN = '' where artnr = " & lartnr
                    sSQL = sSQL & " and EAN = '" & sEAN & "' "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Update artikel set EAN2 = '' where artnr = " & lartnr
                    sSQL = sSQL & " and EAN2 = '" & sEAN & "' "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Update artikel set EAN3 = '' where artnr = " & lartnr
                    sSQL = sSQL & " and EAN3 = '" & sEAN & "' "
                    gdBase.Execute sSQL, dbFailOnError
                
                Else
                        
                    sSQL = "Update EANKOML set farbe = '3' where artnr = " & lartnr
                    sSQL = sSQL & " and (farbe = '2' or farbe = '9') "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    'hier wird der Artikel gelöscht
                    'check ob Liefdaten aufs Duplikat übertragen werden sollen
                    
                    If Check2(0).Value = vbChecked Then
                        sSQL = "Select * from eankoml where ean = '" & sEAN & "' "
                        sSQL = sSQL & " and Linr <> " & lLinr
                        Set rsLi = gdBase.OpenRecordset(sSQL)
                        If Not rsLi.EOF Then
                            rsLi.MoveLast
                            If Not rsLi.EOF Then
                                k = rsLi.RecordCount
                                k = k - 1
                                ReDim lDupliArtnr(k)
                                j = 0
                                rsLi.MoveFirst
                                Do While Not rsLi.EOF
                                    
                                    lDupliArtnr(j) = CLng(rsLi!artnr)
        
                                    j = j + 1
                                    rsLi.MoveNext
                                Loop
                            
                                For l = 0 To k
                                    If Not lDupliArtnr(l) = lartnr Then
                                        'gibt es die artlief kombi schon?
                                        sSQL = "Select * from artlief where artnr = " & lDupliArtnr(l)
                                        sSQL = sSQL & " and Linr = " & lLinr
                                        Set rsDupli = gdBase.OpenRecordset(sSQL)
                                        
                                        If rsDupli.EOF Then
                                            sSQL = "Update artlief set artnr = " & lDupliArtnr(l)
                                            sSQL = sSQL & " Where artnr = " & lartnr
                                            sSQL = sSQL & " and Linr = " & lLinr
                                            gdBase.Execute sSQL, dbFailOnError
                                        End If
                                        rsDupli.Close
                                    
                                    End If
                                Next l
                            End If
                        End If
                        rsLi.Close
                    End If
                
                    'check ob Verkaufzahlen aufs Dupli übertragen werden sollen
                    If Check2(1).Value = vbChecked Then
                        sSQL = "Select * from eankoml where ean = '" & sEAN & "' "
                        sSQL = sSQL & " and artnr <> " & lartnr
                        Set rsLi = gdBase.OpenRecordset(sSQL)
                        If Not rsLi.EOF Then
                            rsLi.MoveLast
                            If rsLi.RecordCount = 1 Then
                                rsLi.MoveFirst
                                lDupArtnr = CLng(rsLi!artnr)
                                    
                                sSQL = "Update Kassjour set artnr = " & lDupArtnr
                                sSQL = sSQL & " Where artnr = " & lartnr
                                gdBase.Execute sSQL, dbFailOnError
                                
                                sSQL = "Update UMS_ART set artnr = " & lDupArtnr
                                sSQL = sSQL & " Where artnr = " & lartnr
                                gdBase.Execute sSQL, dbFailOnError
                                
                                sSQL = "Update UMSARTJ set artnr = " & lDupArtnr
                                sSQL = sSQL & " Where artnr = " & lartnr
                                gdBase.Execute sSQL, dbFailOnError
                                
                            ElseIf rsLi.RecordCount > 1 Then
        '                            MsgBox "Es sind zu viele Duplikate vorhanden, Verkaufszahlen können nicht übertragen werden."
                            End If
                        End If
                        rsLi.Close
                    End If
                    
                    Dim lBestand As Long
                    Dim lNeuBestand As Long
                    
                    If Check2(2).Value = vbChecked Then
                    'Bestand übertragen
                    
                        sSQL = "Select * from eankoml where ean = '" & sEAN & "' "
                        sSQL = sSQL & " and artnr <> " & lartnr
                        Set rsLi = gdBase.OpenRecordset(sSQL)
                        If Not rsLi.EOF Then
                            rsLi.MoveLast
                            If rsLi.RecordCount = 1 Then
                                rsLi.MoveFirst
                                lDupArtnr = CLng(rsLi!artnr)
                                
                                lDupArtnr = CLng(rsLi!artnr)
                                lBestand = CLng(rsLi!BESTAND)
                                
                                lNeuBestand = lBestand + CLng(Text1.Text)
                                
                                Bestandsveraenderung rsLi!artnr, lNeuBestand, "EAN Duplikate"
                                
''''                                Bestandsveraenderung rsLi!artnr, MSFlexGrid1.TextMatrix(lrow, 2), "EAN Duplikate"
                                
                                sSQL = "Update EANKOML set farbe = '4' , bestand = " & MSFlexGrid1.TextMatrix(lrow, 2)
                                sSQL = sSQL & " where artnr = " & lartnr
                                sSQL = sSQL & " and farbe = '9' "
                                gdBase.Execute sSQL, dbFailOnError
                                
                            ElseIf rsLi.RecordCount > 1 Then
        '                            MsgBox "Es sind zu viele Duplikate vorhanden, Bestandszahlen können nicht übertragen werden.", vbInformation, "Winkiss Hinweis:"
                            End If
                        End If
                        rsLi.Close
                        
                    End If
                    
                    If Check2(3).Value = vbChecked Then
                        sSQL = "Select * from eankoml where ean = '" & sEAN & "' "
                        sSQL = sSQL & " and artnr <> " & lartnr
                        Set rsLi = gdBase.OpenRecordset(sSQL)
                        If Not rsLi.EOF Then
                            rsLi.MoveLast
                            If rsLi.RecordCount = 1 Then
                                rsLi.MoveFirst
                                lDupArtnr = CLng(rsLi!artnr)
                                
                                Artikelveraenderung rsLi!artnr, MSFlexGrid1.TextMatrix(lrow, 4), "EAN Duplikate", "KVKPR1"
                                    
                                sSQL = "Update EANKOML set farbe = '5' ,KVKPR1 = '" & MSFlexGrid1.TextMatrix(lrow, 4) & "'"
                                sSQL = sSQL & " where artnr = " & lartnr
                                sSQL = sSQL & " and farbe = '9' "
                                gdBase.Execute sSQL, dbFailOnError
                                
                            ElseIf rsLi.RecordCount > 1 Then
        '                            MsgBox "Es sind zu viele Duplikate vorhanden, der Kassenverkaufspreis kann nicht übertragen werden.", vbInformation, "Winkiss Hinweis:"
                            End If
                        End If
                        rsLi.Close
                    End If
                    
                    SicherInArtikelsic lartnr
                 
                    ' aus Artlief
                    sSQL = "Delete from artlief where artnr = " & lartnr
                    sSQL = sSQL & " and LINR = " & lLinr
                    gdBase.Execute sSQL, dbFailOnError
                        
                    ' aus Artikel
                    sSQL = "Delete from artikel where artnr = " & lartnr
                    gdBase.Execute sSQL, dbFailOnError
                    
                    schreibeProtokollgArtikel "Artikel: " & CStr(lartnr) & " " & ErmittleDetails(CStr(lartnr)) & " wurde gelöscht(EAN Duplikate)."
        
                    rsrs.Close: Set rsrs = Nothing
                End If
            
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.TextMatrix(lrow, 1) = "gelöscht"
                For j = 0 To byAnzahlSpalten - 1
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.Row = lrow
                    MSFlexGrid1.CellBackColor = vbRed
                Next j
            Else
            
            End If
        End If
            
    Next i
    
    MSFlexGrid1.Redraw = True
    
    Screen.MousePointer = 0
    
    MSFlexGrid1.Refresh
    
    anzeige "ERFOLG", "Fertig!", Label1(1)


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "alleweißenlöschen"
    Fehler.gsFehlertext = "Im Programmteil EAN DUPLIKATE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub alleOhneVKuOhneBestandlöschen()
On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim j               As Integer
    Dim lartnr          As Long
    Dim lrow            As Long
    Dim rsrs            As DAO.Recordset
    Dim lcount As Long

    MSFlexGrid1.Row = 0
    MSFlexGrid1.Redraw = False

    Screen.MousePointer = 11
    
    
    
    
    
    lcount = 0

    For i = 2 To MSFlexGrid1.Rows - 1
        lrow = i
        MSFlexGrid1.Row = lrow

        

        MSFlexGrid1.Col = 2

        If MSFlexGrid1.TextMatrix(lrow, 2) = "0" And MSFlexGrid1.TextMatrix(lrow, 7) = "0" Then
        
            lartnr = MSFlexGrid1.TextMatrix(lrow, 0)
             
            ' aufdat älter als 30 Tage
            sSQL = "Update artikel set aufdat = " & CLng(DateValue(Now)) - 100
            sSQL = sSQL & " where aufdat is null"
            sSQL = sSQL & " and artnr = " & lartnr
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update artikel set aufdat = " & CLng(DateValue(Now)) - 100
            sSQL = sSQL & " where trim(aufdat) = '' "
            sSQL = sSQL & " and artnr = " & lartnr
            gdBase.Execute sSQL, dbFailOnError
            
            
            sSQL = "Select * from artikel"
            sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 30
            sSQL = sSQL & " and artnr = " & lartnr
    
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                
                'Kann gelöscht werden
                
                ' aus Artlief
                sSQL = "Delete from artlief where artnr = " & lartnr
                gdBase.Execute sSQL, dbFailOnError
    
                ' aus Artikel
                sSQL = "Delete from artikel where artnr = " & lartnr
                gdBase.Execute sSQL, dbFailOnError
    
                If NewTableSuchenDBKombi("ARTEAN_K", gdBase) = True Then
                    sSQL = "Delete from ARTEAN_K where artnr = " & lartnr
                    gdBase.Execute sSQL, dbFailOnError
                End If
    
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.TextMatrix(lrow, 1) = "gelöscht"
                For j = 0 To byAnzahlSpalten - 1
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.Row = lrow
                    MSFlexGrid1.CellBackColor = vbRed
                Next j
                    
                    
                    
                End If
                rsrs.Close: Set rsrs = Nothing
                
                lcount = lcount + 1

                anzeige "normal", lcount & " Artikel wurden gelöscht...", Label1(1)
            
        End If

    Next i

    MSFlexGrid1.Redraw = True

    Screen.MousePointer = 0

    MSFlexGrid1.Refresh

    anzeige "ERFOLG", "Fertig!", Label1(1)


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "alleOhneVKuOhneBestandlöschen"
    Fehler.gsFehlertext = "Im Programmteil EAN DUPLIKATE ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
End Sub

Private Sub Command5_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet    As Integer
    Dim lrow    As Long
    Dim sEAN    As String
    Dim sSQL    As String
    
    Select Case Index
    
        Case Is = 0
            
            If gcArtNrFiliale <> "" Then
            
                lrow = MSFlexGrid1.Row
   
                sEAN = MSFlexGrid1.TextMatrix(lrow, 3)
                gsARTNR = sEAN
                frmWKL84.Show 1
                
            Else
                Label1(1).Caption = "Bitte einen Artikel markieren!"
                Label1(1).Refresh
            End If
            
        Case Is = 4
            Unload frmWKL75
        Case Is = 1
        
        
            'ist die Spalte drin dann löschen
            If SpalteInTabellegefundenNEW("eankoml", "LASTVK", gdBase) = True Then
                sSQL = "alter table eankoml drop column LASTVK"
                gdBase.Execute sSQL, dbFailOnError
            End If
        
        
            reportbildschirm "", "aWKL33a"
        Case Is = 2
            If gcArtNrFiliale <> "" Then
                frmWKLam.Show 1
            Else
                Label1(1).Caption = "Bitte einen Artikel markieren!"
                Label1(1).Refresh
            End If
        Case Is = 3
            If gcArtNrFiliale <> "" Then
                MSFlexGrid1_DblClick
            Else
                Label1(1).Caption = "Bitte einen Artikel markieren!"
                Label1(1).Refresh
            End If
        Case 5
            'alle Löschen
            alleOhneVKuOhneBestandlöschen
        
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil EAN DUPLIKATE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    Dim sTmp As String

    Check2(0).Value = vbChecked
    Check2(1).Value = vbChecked
    Check2(2).Value = vbChecked
    Check2(3).Value = vbChecked
    
    positionierenwkl75
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    If NewTableSuchenDBKombi("eankoml", gdBase) = True Then
        If SpalteInTabellegefundenNEW("eankoml", "LASTVK", gdBase) = False Then
            anzeigenDupliEan
        Else
            anzeigenDupliEan_mitVK
        End If
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    If NewTableSuchenDBKombi("Artikelsic", gdBase) = False Then
        CreateArtikelsic
    End If
    
     If NewTableSuchenDBKombi("Artliefsic", gdBase) = False Then
        CreateArtliefsic
    End If
    
    
    If SpalteInTabellegefundenNEW("Artikelsic", "DelPcname", gdBase) = False Then
        SpalteAnfuegenNEW "Artikelsic", "DelPcname", "Text(30)", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("Artikelsic", "DelDATE", gdBase) = False Then
        SpalteAnfuegenNEW "Artikelsic", "DelDATE", "DATETIME", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("Artikelsic", "DelTIME", gdBase) = False Then
        SpalteAnfuegenNEW "Artikelsic", "DelTIME", "Text(10)", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("Artliefsic", "DelPcname", gdBase) = False Then
        SpalteAnfuegenNEW "Artliefsic", "DelPcname", "Text(30)", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("Artliefsic", "DelDATE", gdBase) = False Then
        SpalteAnfuegenNEW "Artliefsic", "DelDATE", "DATETIME", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("Artliefsic", "DelTIME", gdBase) = False Then
        SpalteAnfuegenNEW "Artliefsic", "DelTIME", "Text(10)", gdBase
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil EAN DUPLIKATE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub MSFlexGrid1_SelChange()
On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim ctmp       As String
    Dim cSQL       As String
    Dim rsrs       As Recordset
    lrow = MSFlexGrid1.Row

    gcArtNrFiliale = ""
    If lrow <= 1 Then
        
    Else
        If MSFlexGrid1.TextMatrix(lrow, 1) = "gelöscht" Then
            Exit Sub
        Else
            gcArtNrFiliale = MSFlexGrid1.TextMatrix(lrow, 0)
            Text1.Text = MSFlexGrid1.TextMatrix(lrow, 2)
            Text2.Text = MSFlexGrid1.TextMatrix(lrow, 4)
            
            cSQL = "Select Notizen from Artikel where ARTNR = " & gcArtNrFiliale
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then

                ctmp = ""
                If Not IsNull(rsrs!NOTIZEN) Then
                    ctmp = rsrs!NOTIZEN
                End If
                Label1(8).Caption = "Notizen: " & ctmp
                Label1(8).Refresh
            End If
            rsrs.Close: Set rsrs = Nothing
    
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = glSelBack1
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890," & Chr$(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
'    Dim ctmp       As String
'    Dim cSQL       As String
'    Dim rsrs       As Recordset
    lrow = MSFlexGrid1.Row

    gcArtNrFiliale = ""
    If lrow <= 1 Then
        
    Else
        If MSFlexGrid1.TextMatrix(lrow, 1) = "gelöscht" Then
            Exit Sub
        Else
            gcArtNrFiliale = MSFlexGrid1.TextMatrix(lrow, 0)
            Text1.Text = MSFlexGrid1.TextMatrix(lrow, 2)
            Text2.Text = MSFlexGrid1.TextMatrix(lrow, 4)
            
            
            
            
        
'            cSQL = "Select Notizen from Artikel where ARTNR = " & gcArtNrFiliale
'            Set rsrs = gdBase.OpenRecordset(cSQL)
'            If Not rsrs.EOF Then
'
'                ctmp = ""
'                If Not IsNull(rsrs!NOTIZEN) Then
'                    ctmp = rsrs!NOTIZEN
'                End If
'                Label1(8).Caption = "Notizen: " & ctmp
'                Label1(8).Refresh
'            End If
'            rsrs.Close: Set rsrs = Nothing
    
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionierenwkl33"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    
    
    
    If KeyCode = vbKeyF2 Then
    
        lrow = MSFlexGrid1.Row
        gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
        If gsARTNR <> "" Then

            frmWKL10.Show 1
            Me.Refresh
            Screen.MousePointer = 11
            
            MSFlexGrid1.TopRow = lrow
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
            
            Screen.MousePointer = 0
        End If
        gsARTNR = ""
    End If
    
    
    If KeyCode = vbKeyF4 Then
    
        gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
        If gsARTNR <> "" Then
            frmWKL63.Show 1
            Me.Refresh
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
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1

End Sub
Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR

    Dim lRow2   As Long
    Dim sSQL    As String
    Dim lartnr  As Long
    Dim lrow    As Long
    Dim iRet    As Integer
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim l       As Integer
    Dim lcol    As Long
    Dim sEAN    As String
    Dim rsrs    As Recordset
    Dim rsLi    As Recordset
    Dim rsArtLi As Recordset
    Dim rsDupli As Recordset
    Dim lLinr   As Long
    Dim lDupliArtnr() As Long
    Dim lDupArtnr As Long
    
    lrow = MSFlexGrid1.Row
    If lrow <= 1 Then
    
        lcol = MSFlexGrid1.Col
        MSFlexGrid1.Col = lcol
        If lcol = 1 Then
            MSFlexGrid1.sOrt = 1
        ElseIf lcol = 3 Then
            MSFlexGrid1.sOrt = 1
        Else
            MSFlexGrid1.sOrt = 2
        End If
        
    Else
        If MSFlexGrid1.TextMatrix(lrow, 1) = "gelöscht" Then
            Exit Sub
        Else
            lartnr = MSFlexGrid1.TextMatrix(lrow, 0)
            
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = 0
            
            sEAN = MSFlexGrid1.TextMatrix(lrow, 3)
            lLinr = Val(MSFlexGrid1.TextMatrix(lrow, 4))
            
            If Check2(4).Value = vbChecked Then
                iRet = vbYes
            Else
                iRet = MsgBox("Möchten Sie den Artikel: " & lartnr & " wirklich löschen?", vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            
            End If
            If iRet = vbYes Then
                sSQL = "Select * from eankoml where artnr =" & lartnr
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveLast
                
                    If rsrs.RecordCount > 1 Then
                    
                        'kommt eine ArtNr mehrmals vor, so wird nur die EAN gelöscht
                        
                        sSQL = "Update artikel set EAN = '' where artnr = " & lartnr
                        sSQL = sSQL & " and EAN = '" & sEAN & "' "
                        gdBase.Execute sSQL, dbFailOnError

                        sSQL = "Update artikel set EAN2 = '' where artnr = " & lartnr
                        sSQL = sSQL & " and EAN2 = '" & sEAN & "' "
                        gdBase.Execute sSQL, dbFailOnError

                        sSQL = "Update artikel set EAN3 = '' where artnr = " & lartnr
                        sSQL = sSQL & " and EAN3 = '" & sEAN & "' "
                        gdBase.Execute sSQL, dbFailOnError

                        If NewTableSuchenDBKombi("ARTEAN_K", gdBase) = True Then
                            sSQL = "Update ARTEAN_K set EAN = '' where artnr = " & lartnr
                            sSQL = sSQL & " and EAN = '" & sEAN & "' "
                            gdBase.Execute sSQL, dbFailOnError
                        End If
                    
                    Else
                            
                        sSQL = "Update EANKOML set farbe = '3' where artnr = " & lartnr
                        sSQL = sSQL & " and (farbe = '2' or farbe = '9') "
                        gdBase.Execute sSQL, dbFailOnError
                        
                        'hier wird der Artikel gelöscht
                        'check ob Liefdaten aufs Duplikat übertragen werden sollen
                        
                        If Check2(0).Value = vbChecked Then
                            sSQL = "Select * from eankoml where ean = '" & sEAN & "' "
                            sSQL = sSQL & " and Linr <> " & lLinr
                            Set rsLi = gdBase.OpenRecordset(sSQL)
                            If Not rsLi.EOF Then
                                rsLi.MoveLast
                                If Not rsLi.EOF Then
                                    k = rsLi.RecordCount
                                    k = k - 1
                                    ReDim lDupliArtnr(k)
                                    j = 0
                                    rsLi.MoveFirst
                                    Do While Not rsLi.EOF
                                        
                                        lDupliArtnr(j) = CLng(rsLi!artnr)
        
                                        j = j + 1
                                        rsLi.MoveNext
                                    Loop
                                
                                    For l = 0 To k
                                        If Not lDupliArtnr(l) = lartnr Then
                                            'gibt es die artlief kombi schon?
                                            sSQL = "Select * from artlief where artnr = " & lDupliArtnr(l)
                                            sSQL = sSQL & " and Linr = " & lLinr
                                            Set rsDupli = gdBase.OpenRecordset(sSQL)
                                            
                                            If rsDupli.EOF Then
                                                sSQL = "Update artlief set artnr = " & lDupliArtnr(l)
                                                sSQL = sSQL & " Where artnr = " & lartnr
                                                sSQL = sSQL & " and Linr = " & lLinr
                                                gdBase.Execute sSQL, dbFailOnError
                                            End If
                                            rsDupli.Close
                                        
                                        End If
                                    Next l
                                End If
                            End If
                            rsLi.Close
                        End If
                    
                        'check ob Verkaufzahlen aufs Dupli übertragen werden sollen
                        If Check2(1).Value = vbChecked Then
                            sSQL = "Select * from eankoml where ean = '" & sEAN & "' "
                            sSQL = sSQL & " and artnr <> " & lartnr
                            Set rsLi = gdBase.OpenRecordset(sSQL)
                            If Not rsLi.EOF Then
                                rsLi.MoveLast
                                If rsLi.RecordCount = 1 Then
                                    rsLi.MoveFirst
                                    lDupArtnr = CLng(rsLi!artnr)
                                        
                                    sSQL = "Update Kassjour set artnr = " & lDupArtnr
                                    sSQL = sSQL & " Where artnr = " & lartnr
                                    gdBase.Execute sSQL, dbFailOnError
                                    
                                    sSQL = "Update UMS_ART set artnr = " & lDupArtnr
                                    sSQL = sSQL & " Where artnr = " & lartnr
                                    gdBase.Execute sSQL, dbFailOnError
                                    
                                    sSQL = "Update UMSARTJ set artnr = " & lDupArtnr
                                    sSQL = sSQL & " Where artnr = " & lartnr
                                    gdBase.Execute sSQL, dbFailOnError
                                    
                                ElseIf rsLi.RecordCount > 1 Then
                                    MsgBox "Es sind zu viele Duplikate vorhanden, Verkaufszahlen können nicht übertragen werden."
                                End If
                            End If
                            rsLi.Close
                        End If
                        
                        Dim lBestand As Long
                        Dim lNeuBestand As Long
                        
                        If Check2(2).Value = vbChecked Then
                        'Bestand übertragen
                        
                            sSQL = "Select * from eankoml where ean = '" & sEAN & "' "
                            sSQL = sSQL & " and artnr <> " & lartnr
                            Set rsLi = gdBase.OpenRecordset(sSQL)
                            If Not rsLi.EOF Then
                                rsLi.MoveLast
                                If rsLi.RecordCount = 1 Then
                                    rsLi.MoveFirst
                                    lDupArtnr = CLng(rsLi!artnr)
                                    lBestand = CLng(rsLi!BESTAND)
                                    
                                    lNeuBestand = lBestand + CLng(Text1.Text)
                                    
                                    Bestandsveraenderung rsLi!artnr, lNeuBestand, "EAN Duplikate"
                                    
                                    sSQL = "Update EANKOML set farbe = '4' , bestand = " & Text1.Text
                                    sSQL = sSQL & " where artnr = " & lartnr
                                    sSQL = sSQL & " and farbe = '9' "
                                    gdBase.Execute sSQL, dbFailOnError
                                    
                                ElseIf rsLi.RecordCount > 1 Then
                                    MsgBox "Es sind zu viele Duplikate vorhanden, Bestandszahlen können nicht übertragen werden.", vbInformation, "Winkiss Hinweis:"
                                End If
                            End If
                            rsLi.Close
                            
                        End If
                        
                        If Check2(3).Value = vbChecked Then
                            sSQL = "Select * from eankoml where ean = '" & sEAN & "' "
                            sSQL = sSQL & " and artnr <> " & lartnr
                            Set rsLi = gdBase.OpenRecordset(sSQL)
                            If Not rsLi.EOF Then
                                rsLi.MoveLast
                                If rsLi.RecordCount = 1 Then
                                    rsLi.MoveFirst
                                    lDupArtnr = CLng(rsLi!artnr)
                                    
                                    Artikelveraenderung rsLi!artnr, Text2.Text, "EAN Duplikate", "KVKPR1"
                                        
                                    sSQL = "Update EANKOML set farbe = '5' ,KVKPR1 = '" & Text2.Text & "'"
                                    sSQL = sSQL & " where artnr = " & lartnr
                                    sSQL = sSQL & " and farbe = '9' "
                                    gdBase.Execute sSQL, dbFailOnError
                                    
                                ElseIf rsLi.RecordCount > 1 Then
                                    MsgBox "Es sind zu viele Duplikate vorhanden, der Kassenverkaufspreis kann nicht übertragen werden.", vbInformation, "Winkiss Hinweis:"
                                End If
                            End If
                            rsLi.Close
                        End If
                        
                        SicherInArtikelsic lartnr
                     
                        ' aus Artlief
                        sSQL = "Delete from artlief where artnr = " & lartnr
                        sSQL = sSQL & " and LINR = " & lLinr
                        gdBase.Execute sSQL, dbFailOnError
                            
                        ' aus Artikel
                        sSQL = "Delete from artikel where artnr = " & lartnr
                        gdBase.Execute sSQL, dbFailOnError
                        
                        ' aus ArtEAN_K
                        
                        If NewTableSuchenDBKombi("ARTEAN_K", gdBase) = True Then
                            sSQL = "Delete from ARTEAN_K where artnr = " & lartnr
                            gdBase.Execute sSQL, dbFailOnError
                        End If
                        
                        
                        
                        
                        
                        schreibeProtokollgArtikel "Artikel: " & CStr(lartnr) & " " & ErmittleDetails(CStr(lartnr)) & " wurde gelöscht(EAN Duplikate)."
    
                        rsrs.Close: Set rsrs = Nothing
                    End If
                
                    MSFlexGrid1.Row = lrow
                    MSFlexGrid1.TextMatrix(lrow, 1) = "gelöscht"
                    For i = 0 To byAnzahlSpalten - 1
                        MSFlexGrid1.Col = i
                        MSFlexGrid1.Row = lrow
                        MSFlexGrid1.CellBackColor = vbRed
                    Next i
                Else
                
                End If
            End If
        End If
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
    
    Case Is = vbKeyReturn
            MSFlexGrid1_Click
            Command5_Click 2
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub positionierenwkl75()
On Error GoTo LOKAL_ERROR

    Frame1.Height = 7575
    Frame1.Width = 11775
    Frame1.Top = 960
    Frame1.Left = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionierenwkl75"
    Fehler.gsFehlertext = "Im Programmteil EAN DUPLIKATE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub formatgrid()
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer
    
    With MSFlexGrid1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 7
        byAnzahlSpalten = .Cols
         ReDim aBreite(.Cols)
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .Text = "Artnr"
        
        .Col = 1
        .Text = "Artikelbezeichnung"
        
        .Col = 2
        .Text = "Bestand"
        
        .Col = 3
        .Text = "EAN"
        
        .Col = 4
        .Text = "Kassenpreis"
        
        .Col = 5
        .Text = "LiNr"
        
        .Col = 6
        .Text = "Lieferantenbezeichnung"
        
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "formatgrid"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub formatgrid_mitVK()
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer
    
    With MSFlexGrid1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 8
        byAnzahlSpalten = .Cols
         ReDim aBreite(.Cols)
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .Text = "Artnr"
        
        .Col = 1
        .Text = "Artikelbezeichnung"
        
        .Col = 2
        .Text = "Bestand"
        
        .Col = 3
        .Text = "EAN"
        
        .Col = 4
        .Text = "Kassenpreis"
        
        .Col = 5
        .Text = "LiNr"
        
        .Col = 6
        .Text = "Lieferantenbezeichnung"
        
        .Col = 7
        .Text = "letzer Verkauf"
        
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "formatgrid_mitVK"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub anzeigenDupliEan()
    On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim j As Integer
    Dim cPfad As String
    Dim lrow As Long
    Dim rsrs  As Recordset
    Dim lWert As Long
    Dim sWert As String
    Dim dWert As Double
    
    Dim bAlleWhite As Boolean
    
    bAlleWhite = False
    
    MSFlexGrid1.Redraw = False
    
    lrow = 1
    
   
    
    
    
'    Set rsrs = gdBase.OpenRecordset(sSQL)
    Set rsrs = gdBase.OpenRecordset("eankoml", dbOpenTable)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        
        anzeige "ERFOLG", "Alles Prima! Keine EAN Duplikate gefunden.", Label1(1)
        
        Frame2.Visible = False
        Command5(0).Visible = False
        Command5(1).Visible = False
        Command5(2).Visible = False
        Check2(4).Visible = False
        Check2(5).Visible = False
        Label1(0).Visible = False
        
        Frame1.Visible = True
        
        Exit Sub
    Else
        formatgrid
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            lrow = lrow + 1
            MSFlexGrid1.Rows = lrow + 1
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = 0
            
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = 0
            End If
            MSFlexGrid1.Text = lWert
            
            If Not IsNull(rsrs!ARTIKELBEZEICHNUNG) Then
                sWert = rsrs!ARTIKELBEZEICHNUNG
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 1
'            MSFlexGrid1.CellAlignment = Left
            MSFlexGrid1.Text = sWert
            
            If Not IsNull(rsrs!BESTAND) Then
                lWert = rsrs!BESTAND
            Else
                lWert = 0
            End If
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = lWert
            
            If Not IsNull(rsrs!EAN) Then
                sWert = (rsrs!EAN)
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 3
            
            MSFlexGrid1.Text = sWert
            
            If Not IsNull(rsrs!KVKPR1) Then
                dWert = rsrs!KVKPR1
            Else
                dWert = 0
            End If
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = Format$(dWert, "######0.00")
            
            If Not IsNull(rsrs!linr) Then
                sWert = (rsrs!linr)
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 5
            
            MSFlexGrid1.Text = sWert
            
            If Not IsNull(rsrs!linbez) Then
                sWert = (rsrs!linbez)
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 6
            
            MSFlexGrid1.Text = sWert
            
             If Not IsNull(rsrs!farbe) Then
                sWert = (rsrs!farbe)
            Else
                sWert = "1"
            End If
            
            If sWert = "2" Then

                For j = 1 To byAnzahlSpalten - 1
                    Label1(2).Visible = True
                    Label1(3).Visible = True
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.CellBackColor = vbGreen
                    MSFlexGrid1.CellForeColor = vbBlack
                    
                Next j
                bAlleWhite = True
            ElseIf sWert = "9" Then
                For j = 1 To byAnzahlSpalten - 1
                    Label1(4).Visible = True
                    Label1(5).Visible = True
                    Label1(6).Visible = True
                    Label1(7).Visible = True
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.CellBackColor = vbYellow
                    MSFlexGrid1.CellForeColor = vbBlack
                    
                Next j
                bAlleWhite = True
            Else
                For j = 1 To byAnzahlSpalten - 1
                    
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.CellBackColor = vbWhite
                    MSFlexGrid1.CellForeColor = vbBlack
                Next j
            
            End If
            
            
            
            rsrs.MoveNext
        Loop
        
        MSFlexGrid1.RowHeight(1) = 0
        
        Label1(1).Caption = rsrs.RecordCount & " mehrfach angelegte Artikel ermittelt (Übereinstimmung im EAN - CODE)."
        Label1(1).Refresh
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Frame1.Visible = True
    
    lrow = 0
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    

    If bAlleWhite Then
        Check2(5).Visible = True
    Else
        Check2(5).Visible = False
    End If
    
    'alle Löschen ausblenden
    Command5(5).Visible = False
    Label1(9).Visible = False
    'ende

    MSFlexGrid1.Redraw = True
    MSFlexGrid1.Visible = True
    Command5(1).Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "anzeigenDupliEan"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub anzeigenDupliEan_mitVK()
    On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim j As Integer
    Dim cPfad As String
    Dim lrow As Long
    Dim rsrs  As Recordset
    Dim lWert As Long
    Dim sWert As String
    Dim dWert As Double
    
    Dim bAlleWhite As Boolean
    
    bAlleWhite = False
    
    MSFlexGrid1.Redraw = False
    
    lrow = 1

    Set rsrs = gdBase.OpenRecordset("eankoml", dbOpenTable)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        
        anzeige "ERFOLG", "Alles Prima! Keine EAN Duplikate gefunden.", Label1(1)
        
        Frame2.Visible = False
        Command5(0).Visible = False
        Command5(1).Visible = False
        Command5(2).Visible = False
        Check2(4).Visible = False
        Check2(5).Visible = False
        Label1(0).Visible = False
        
        Frame1.Visible = True
        
        Exit Sub
    Else
        formatgrid_mitVK
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            lrow = lrow + 1
            MSFlexGrid1.Rows = lrow + 1
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = 0
            
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = 0
            End If
            MSFlexGrid1.Text = lWert
            
            If Not IsNull(rsrs!ARTIKELBEZEICHNUNG) Then
                sWert = rsrs!ARTIKELBEZEICHNUNG
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 1
'            MSFlexGrid1.CellAlignment = Left
            MSFlexGrid1.Text = sWert
            
            If Not IsNull(rsrs!BESTAND) Then
                lWert = rsrs!BESTAND
            Else
                lWert = 0
            End If
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = lWert
            
            If Not IsNull(rsrs!EAN) Then
                sWert = (rsrs!EAN)
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 3
            
            MSFlexGrid1.Text = sWert
            
            If Not IsNull(rsrs!KVKPR1) Then
                dWert = rsrs!KVKPR1
            Else
                dWert = 0
            End If
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = Format$(dWert, "######0.00")
            
            If Not IsNull(rsrs!linr) Then
                sWert = (rsrs!linr)
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 5
            
            MSFlexGrid1.Text = sWert
            
            If Not IsNull(rsrs!linbez) Then
                sWert = (rsrs!linbez)
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 6
            
            MSFlexGrid1.Text = sWert
            
            
            
            If Not IsNull(rsrs!lastvk) Then
                sWert = (rsrs!lastvk)
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 7
            
            MSFlexGrid1.Text = sWert
            
            
            
            
            
            
            
            
            
            
             If Not IsNull(rsrs!farbe) Then
                sWert = (rsrs!farbe)
            Else
                sWert = "1"
            End If
            
            If sWert = "2" Then

                For j = 1 To byAnzahlSpalten - 1
                    Label1(2).Visible = True
                    Label1(3).Visible = True
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.CellBackColor = vbGreen
                    MSFlexGrid1.CellForeColor = vbBlack
                    
                Next j
                bAlleWhite = True
            ElseIf sWert = "9" Then
                For j = 1 To byAnzahlSpalten - 1
                    Label1(4).Visible = True
                    Label1(5).Visible = True
                    Label1(6).Visible = True
                    Label1(7).Visible = True
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.CellBackColor = vbYellow
                    MSFlexGrid1.CellForeColor = vbBlack
                    
                Next j
                bAlleWhite = True
            Else
                For j = 1 To byAnzahlSpalten - 1
                    
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.CellBackColor = vbWhite
                    MSFlexGrid1.CellForeColor = vbBlack
                Next j
            
            End If
            
            
            
            rsrs.MoveNext
        Loop
        
        MSFlexGrid1.RowHeight(1) = 0
        
        Label1(1).Caption = rsrs.RecordCount & " mehrfach angelegte Artikel ermittelt (Übereinstimmung im EAN - CODE)."
        Label1(1).Refresh
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Frame1.Visible = True
    
    lrow = 0
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    

    If bAlleWhite Then
        Check2(5).Visible = True
    Else
        Check2(5).Visible = False
    End If

    MSFlexGrid1.Redraw = True
    MSFlexGrid1.Visible = True
    Command5(1).Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "anzeigenDupliEan_mitVK"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
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
