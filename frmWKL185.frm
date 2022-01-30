VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL185 
   BackColor       =   &H00C0C000&
   Caption         =   "neue Rewe - Artikeldaten"
   ClientHeight    =   8610
   ClientLeft      =   1215
   ClientTop       =   1590
   ClientWidth     =   11910
   Icon            =   "frmWKL185.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chk_Pflichtpreisanpassungen_vornehmen 
      Caption         =   "Pflichtpreisanpassungen vornehmen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   22
      Top             =   5400
      Value           =   1  'Aktiviert
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11160
      Top             =   240
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   6
      Left            =   6840
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
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
      ButtonStyle     =   2
      Caption         =   "Etiketten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
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
      ButtonStyle     =   2
      Caption         =   "Etiketten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
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
      ButtonStyle     =   2
      Caption         =   "Etiketten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   16
      Top             =   7080
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
      Caption         =   "Protokoll"
      Enabled         =   0   'False
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   10320
      Pattern         =   "MASTER!.*"
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Top             =   7680
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
      Caption         =   "Schlieﬂen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   0
      Top             =   6480
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
      Caption         =   "Einlesen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   13
      Left            =   4080
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "VK Pflichtpreisanpassung"
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
      Index           =   12
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "n‰here Informationen hier: (bitte anklicken)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   7920
      MouseIcon       =   "frmWKL185.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   15
      ToolTipText     =   "hier alle Neuigkeiten lesen"
      Top             =   5880
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   10
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Artikel insgesamt"
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
      Index           =   9
      Left            =   1440
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   8
      Left            =   4080
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   7
      Left            =   4080
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   6
      Left            =   4080
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "EK Preis‰nderung"
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
      Index           =   5
      Left            =   1440
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Auslistung / Ex "
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
      Index           =   3
      Left            =   1440
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "neue Artikel"
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
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Mˆchten Sie diese ¸bernehmen, so klicken Sie auf ""Einlesen""."
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
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neue Rewe-Artikeldaten stehen bereit. "
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
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   9015
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "neue Rewe - Artikeldaten"
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
      TabIndex        =   3
      Top             =   120
      Width           =   11535
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
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL185"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim glartv      As Long
Dim glartb      As Long
Dim iSec        As Integer
Private Sub ReweDatenEinlesen(sPfad As String, sDatei As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsRS    As Recordset
    
    loeschNEW "REWE_AGN", gdApp
    CreateTableT2 "REWE_AGN", gdApp

    'vorbereitung der Importtabelle
    '1. erst lˆschen

    loeschNEW "MEISTER", gdApp
    CreateTable "MEISTER", gdApp

    Dim iFileNr As Integer
    Dim cSatz1 As String
    Dim lPos As Long
    Dim cEinzelsatz As String
    Dim lLinr As Long
    Dim lfnr1 As Long
    Dim cBezeich As String
    Dim rsagn As Recordset
    Dim lPosEnde As Long
    Dim lLenfil As Long
    Dim lposSemi As Long
    Dim lposSemiEnde As Long
    Dim cWert As String
    Dim cArtbez_1 As String
    Dim cArtbez_2 As String
    Dim cInhalt As String
    Dim cInhaltBez As String
    Dim cInhaltsangabegesamt As String
    Dim cAgn As String
    Dim cAGNBEZEICH As String
    Dim cWarengruppe As String
    
    Label1(2).Visible = True
    Label1(3).Visible = True
    Label1(5).Visible = True
    Label1(9).Visible = True
    Label1(6).Visible = True
    Label1(7).Visible = True
    Label1(8).Visible = True
    Label1(10).Visible = True
    Label1(12).Visible = True
    Label1(13).Visible = True
    
    Command5(3).Enabled = False
    Command5(5).Enabled = False
    Command5(6).Enabled = False
    
    lfnr1 = 0
    lPos = 1
    lPosEnde = 1
    lposSemiEnde = 1

    Set rsRS = gdApp.OpenRecordset("MEISTER")
    Set rsagn = gdApp.OpenRecordset("REWE_AGN")

    iFileNr = FreeFile
    Open sPfad & "\" & sDatei For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then

        cSatz1 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz1


        lLenfil = Len(cSatz1)

        lPosEnde = InStr(lPos, cSatz1, vbCrLf)
        lPos = lPos + lPosEnde - lPos + 2 'Kopfzeile ¸berspringen

        Do
            lPosEnde = InStr(lPos, cSatz1, vbCrLf)

            cEinzelsatz = Mid(cSatz1, lPos, lPosEnde - lPos)
            
            lPos = lPos + lPosEnde - lPos + 2

            lposSemi = 1

            rsRS.AddNew
            rsagn.AddNew
            lfnr1 = lfnr1 + 1
            rsRS!lfnr = lfnr1

            'Libesnr
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsRS!LIBESNR = cWert

            'Liefnr ohne Bedeutung
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            'Liefname = Notizen
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsRS!NOTIZEN = cWert

            'WGRU = AGN
            cWert = ""
            cAgn = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            cAgn = cWert
            rsRS!AGN = CLng(cAgn)
            rsagn!RAGN = CLng(cAgn)

            'WGRU_BEZ
            cWert = ""
            cAGNBEZEICH = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            cAGNBEZEICH = Left(cWert, 30)
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "Ω", "ƒ")
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "!", "‹")
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "\", "÷")
            rsagn!RAGTEXT = cAGNBEZEICH

            'Artbez_1
            cWert = ""
            cArtbez_1 = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            cArtbez_1 = cWert

            'Artbez_2
            cWert = ""
            cArtbez_2 = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            cArtbez_2 = cWert

            'ME_TEXT ohne Bedeutung
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            'EK = EKPR
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            If IsNumeric(cWert) Then
                rsRS!lekpr = cWert
            Else
                rsRS!lekpr = 0
            End If

            'UVP = VKPR
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            If IsNumeric(cWert) Then
                rsRS!vkpr = cWert
            Else
                rsRS!vkpr = 0
            End If

            'EAN
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsRS!EAN = cWert

            'MWST
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            If cWert <> "" Then
                If cWert = 1 Then
                    rsRS!MWST = "V"
                ElseIf cWert = 2 Then
                    rsRS!MWST = "E"
                End If
            End If

            'Einh = Minmen
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsRS!MINMEN = cWert

            'Inhalt
            cWert = ""
            cInhalt = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            cInhalt = Trim(cWert)
            
            If cInhalt <> "" Then
                cInhalt = SwapStr(cInhalt, ",", "")
                cInhalt = CStr(CLng(cInhalt) / 10000)
                cInhalt = Format(cInhalt, "####0.00")
                
                If Right(cInhalt, 3) = ",00" Then
                    cInhalt = Left(cInhalt, Len(cInhalt) - 3)
                End If
                
                If Right(cInhalt, 2) = ",0" Then
                    cInhalt = Left(cInhalt, Len(cInhalt) - 2)
                End If
                
                If Right(cInhalt, 1) = "0" And InStr(1, cInhalt, ",") > 0 Then
                    cInhalt = Left(cInhalt, Len(cInhalt) - 1)
                End If
            End If
            
            If cInhalt <> "" Then
                If cInhalt = "0" Then
                    cInhalt = ""
                Else
                    rsRS!INHALT = cInhalt
                End If
            End If

            'Einheit des Inhalts
            cWert = ""
            cInhaltBez = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            cInhaltBez = Trim(cWert)
            rsRS!INHALTBEZ = cInhaltBez
            
            'Auslistung
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            If UCase(cWert) = "X" Then
                rsRS!RKZ = "J"
                Label1(7).Caption = CInt(Label1(7).Caption) + 1
                Label1(7).Refresh
            Else
                rsRS!RKZ = "N"
            End If
            
            'UVP-Pflicht
            cWert = ""
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            If chk_Pflichtpreisanpassungen_vornehmen.Value = vbChecked Then
                If UCase(cWert) = "X" Then
                    rsRS!Status = "J"
                Else
                    rsRS!Status = "N"
                End If
            Else
                rsRS!Status = "N"
            End If
            
            
            
            'EAN2
            cWert = ""
            
'            MsgBox Len(cEinzelsatz)



            'zeig mal den Rest
            
            Dim sRest As String
            
            If Len(cEinzelsatz) > lposSemiEnde Then
                sRest = Mid(cEinzelsatz, lposSemi, Len(cEinzelsatz) - lposSemi + 1)
                cWert = Trim(sRest)
            End If


'            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
'
'
'            If lposSemiEnde > 0 Then
'                cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
'                lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
'
'            End If
            
            If cWert <> "" Then
                cWert = Val(cWert)
                If Len(cWert) = 11 Then
                    cWert = "0" & cWert
                End If
            End If
            
            rsRS!EAN2 = cWert
            rsRS!EAN3 = ""
        
            rsRS!GRUNDPREIS = "J"

            'Bezeichnung aufbauen
            cArtbez_2 = Trim(cArtbez_2)
            cArtbez_2 = SwapStr(cArtbez_2, "  ", " ")

            If Trim(cArtbez_2) <> "" Then
                If Len(cArtbez_2) >= 8 Then
                    cArtbez_2 = Left(cArtbez_2, 8)
                End If
            End If
            
            cBezeich = ""

            If cArtbez_2 <> "" Then
                cBezeich = cArtbez_2
            End If

            cArtbez_1 = Trim(cArtbez_1)
            cArtbez_1 = SwapStr(cArtbez_1, "  ", " ")

            cInhalt = Trim(cInhalt)
            cInhaltBez = Trim(cInhaltBez)

            cInhaltsangabegesamt = cInhalt & cInhaltBez

            If cArtbez_1 <> "" Then
                If cInhaltsangabegesamt <> "" Then

                    cBezeich = cBezeich & " " & cArtbez_1
                    cBezeich = Left(cBezeich, 34 - Len(cInhaltsangabegesamt))
                    cBezeich = cBezeich & " " & cInhaltsangabegesamt

                Else
                    cBezeich = cBezeich & " " & cArtbez_1
                    cBezeich = Left(cBezeich, 35)
                End If

            End If

            rsRS!BEZEICH = cBezeich

            cWarengruppe = ""

            Select Case Left(cAgn, 2)
                Case "09", "13", "15", "16", "21", "26", "27", "28", "31", "32", "33", "34", "35", "36"
                    cWarengruppe = "Lebensmittel"
                Case "29", "68"
                    cWarengruppe = "Babynahrung/-bedarf"
                Case "18", "72"
                    cWarengruppe = "Gartenbedarf"
                Case "30", "53", "54", "56", "57", "58", "59", "61", "62", "74", "75", "76", "77", "82"
                    cWarengruppe = "Drogeriewaren"
                Case "38", "39", "40", "64"
                    cWarengruppe = "S¸ﬂwaren, Chips"
                Case "42", "43", "44", "45", "46", "47", "48"
                    cWarengruppe = "Getr‰nke"
                Case "49", "50", "51"
                    cWarengruppe = "Kaffee/Tee/Kakao/Tabak"
                Case "65"
                    cWarengruppe = "Tiernahrung/-bedarf"
                Case "95", "96", "98", "99"
                    cWarengruppe = "Sonstiges"
                Case Else
                    cWarengruppe = "unbekannt"
            End Select

            rsRS!MARKE = cWarengruppe

            rsRS!LPZ = 1
            rsRS!linr = "400001"
            rsRS!GEFUEHRT = "J"


            rsagn.Update
            rsRS.Update
            
            Label1(10).Caption = CInt(Label1(10).Caption) + 1
            Label1(10).Refresh

        Loop While lLenfil >= lPos
    End If

    Close iFileNr
    rsRS.Close: Set rsRS = Nothing
    rsagn.Close: Set rsagn = Nothing

    '****************

    '****************
    'hier Reweagn's auf neue pr¸fen

    loeschNEW "REWE_AGN_DIS", gdApp

    sSQL = "Select ragn as agn, ragtext as agtext into REWE_AGN_DIS from REWE_AGN group by ragn,ragtext "
    gdApp.Execute sSQL, dbFailOnError

    loeschNEW "REWE_AGN", gdApp

    loeschNEW "REWE_AGN_DIS", gdBase
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "REWE_AGN_DIS"

    sSQL = "Delete from REWE_AGN_DIS where REWE_AGN_DIS.agn in (Select agn from agndbf)"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into agndbf select agn,agtext from REWE_AGN_DIS"
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "REWE_AGN_DIS", gdBase

    'Ende hier Reweagn's auf neue pr¸fen
    
    anzeige "normal", "EAN Duplikate werden entfernt...", Label1(4)

    '6. EAN - Duplikats¸berpr¸fung in der Importtabelle Anzahl ermitteln
    ErmittlungReweDuplisPlusDel
    
    
    
    '
    
    
    
    
    
        
    sSQL = "Delete from Meister where LEKPR = 0 "
    gdApp.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Die Artikeldaten werden ¸berpr¸ft...", Label1(4)
   
    '7. diverse Feld¸berpr¸fungen vornehmen
    feldcheckRewe
    anzeige "normal", "Die Sortiments¸bersicht wird erstellt...", Label1(4)

    '8. Tabelle IMPORTPRI zur Datenbank kopieren
    loeschNEW "ImportpriREWE", gdBase
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "ImportpriREWE"

    sSQL = "Create Index ARTNR on ImportpriREWE (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update ImportpriREWE set marke ='' "
    gdBase.Execute sSQL, dbFailOnError
    
    FormatiereBildschirmdatenREWE
    
    'jetzt ¸bernehmen
    Uebernahme_Rewe_Delta sDatei
    
    If NewTableSuchenDBKombi("ImportpriREWE_dauer", gdBase) = False Then
        sSQL = "select * into ImportpriREWE_dauer from ImportpriREWE  "
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Insert into ImportpriREWE_dauer select * from ImportpriREWE  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Kill sPfad & "\" & sDatei
    
    Command5(2).Enabled = True
    
    Command5(3).Enabled = True
    Command5(5).Enabled = True
    Command5(6).Enabled = True
    
    anzeige "normal", "Fertig! Artikel¸bernahme beendet, Protokoll und Etiketten ausdrucken!", Label1(4)
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ReweDatenEinlesen"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
    
End Sub
Private Sub Uebernahme_Rewe_Delta(sDatei As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsRS            As DAO.Recordset
    Dim rsArtlief       As DAO.Recordset
    Dim sArtnr          As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen...", Label1(4)
    
    If NewTableSuchenDBKombi("STADAPROREWE", gdBase) = False Then
        CreateTableT2 "STADAPROREWE", gdBase
    End If
    
    
''    hier zwischen ARTEAN_K f¸llen
'    loeschNEW "artean_REWE", gdBase
'
'    cSQL = "Insert into artean_REWE Select artnr, val(ean2) as EAN from ImportpriREWE  "
'    gdBase.Execute cSQL, dbFailOnError
'
'
'    'ende hier zwischen ARTEAN_K f¸llen
    
    
    
    
    'Alle EANS auffangen************************************************

    loeschNEW "artean_BU", gdBase

    cSQL = "Create Table artean_BU  "
    cSQL = cSQL & "( ARTNR int"
    cSQL = cSQL & ", EANCH varchar(13)"
    cSQL = cSQL & ", OTHERARTNR1 int"
    cSQL = cSQL & ", OTHERARTNR1_Bestand int"
    
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError


    cSQL = "Insert into artean_BU Select artnr, val(ean) as EANCH from ImportpriREWE  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into artean_BU Select artnr, val(ean2) as EANCH from ImportpriREWE  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_BU set EANCH = '0' & EANCH  "
    cSQL = cSQL & " where len(EANCH)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_BU set OTHERARTNR1 = 0 , OTHERARTNR1_Bestand = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_BU where EANCH = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    loeschNEW "artean_Artikel", gdBase

    cSQL = "Create Table artean_Artikel  "
    cSQL = cSQL & "( ARTNR int"
    cSQL = cSQL & ", EANCH varchar(13)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN = '' where EAN is null  "
    gdBase.Execute cSQL, dbFailOnError


    cSQL = "Insert into artean_Artikel Select artnr, val(ean) as EANCH from Artikel  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN2 = '' where EAN2 is null  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into artean_Artikel Select artnr, val(ean2) as EANCH from Artikel  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN3 = '' where EAN3 is null  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into artean_Artikel Select artnr, val(ean3) as EANCH from Artikel  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    If NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then
        'nich vergessen ARTEAN_K
        cSQL = "Update ARTEAN_K set EAN = '' where EAN is null  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Insert into artean_Artikel Select artnr, val(ean) as EANCH from ARTEAN_K  "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    
    cSQL = "Delete from artean_Artikel where EANCH = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update artean_Artikel set EANCH = '0' & EANCH  "
    cSQL = cSQL & " where len(EANCH)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    
    'mit dieser Information holen wir uns aus der Datenbank alle ARTNR(ungleich), die schon solch eine EAN besitzen
    
    cSQL = "Update artean_BU inner join artean_Artikel on artean_BU.EANCH = artean_Artikel.EANCH "
    cSQL = cSQL & " set artean_BU.OTHERARTNR1 = artean_Artikel.ARTNR "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_BU where Artnr = OTHERARTNR1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_BU where OTHERARTNR1 = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_BU inner join Artikel on artean_BU.OTHERARTNR1 = Artikel.ARTNR "
    cSQL = cSQL & " set artean_BU.OTHERARTNR1_Bestand = Artikel.Bestand "
    gdBase.Execute cSQL, dbFailOnError
    
    
    'nun zur Vollendung
    'bei Otherartnr1 die EAN rausnehmen und von Otherartnr1 auf Artnr den Bestand addieren
    
    cSQL = "Update Artikel inner join artean_BU on Artikel.ARTNR = artean_BU.OTHERARTNR1 and Artikel.EAN = artean_BU.EANCH "
    cSQL = cSQL & " set Artikel.ean = '' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel inner join artean_BU on Artikel.ARTNR = artean_BU.OTHERARTNR1 and Artikel.EAN2 = artean_BU.EANCH "
    cSQL = cSQL & " set Artikel.ean2 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel inner join artean_BU on Artikel.ARTNR = artean_BU.OTHERARTNR1 and Artikel.EAN3 = artean_BU.EANCH "
    cSQL = cSQL & " set Artikel.ean3 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then
        cSQL = "Update ARTEAN_K inner join artean_BU on ARTEAN_K.ARTNR = artean_BU.OTHERARTNR1 and ARTEAN_K.EAN = artean_BU.EANCH "
        cSQL = cSQL & " set ARTEAN_K.ean = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from ARTEAN_K where ean = '' "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    
    'bestand addieren
    
    cSQL = "Update Artikel inner join artean_BU on Artikel.ARTNR = artean_BU.ARTNR  "
    cSQL = cSQL & " set Artikel.Bestand = Artikel.Bestand + artean_BU.OTHERARTNR1_Bestand"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    

    'Ende ********************** Alle EANS auffangen
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Ex Artikel, und Exdatum
    cSQL = "Update Artikel inner join ImportpriREWE i on artikel.artnr = i.artnr "
    cSQL = cSQL & " set "
    cSQL = cSQL & " artikel.LASTDATE = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", artikel.LASTTIME = '" & TimeValue(Now) & "' "
    cSQL = cSQL & ", artikel.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artlief inner join ImportpriREWE i on Artlief.artnr = i.artnr  "
    cSQL = cSQL & " set Artlief.RKZ = 'J'"
    cSQL = cSQL & ", Artlief.EXDAT = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", Artlief.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    cSQL = cSQL & "  and Artlief.linr = 400001 "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(1)...", Label1(4)
    
    
    If Val(Label1(7).Caption) > 0 Then
        Command5(3).Visible = True
        Command5(3).Enabled = False
    End If
    
    'Ex Artikel
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROREWE Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'Auslistung/Ex' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & " , '" & DateValue(Now) & "' as EXDat "
    cSQL = cSQL & " from ImportpriREWE where rkz = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(2)...", Label1(4)
    
    'neue Artikel
    cSQL = "Select * from ImportpriREWE where awm = '98'"
    Set rsRS = gdBase.OpenRecordset(cSQL)
    If Not rsRS.EOF Then
        rsRS.MoveFirst
        Do While Not rsRS.EOF
        
            If Not IsNull(rsRS!artnr) Then
                sArtnr = Trim(rsRS!artnr)
            End If
            
            Sicherheitslˆschen sArtnr 'artlief

            rsRS.MoveNext
            
        Loop
    End If

    rsRS.Close: Set rsRS = Nothing
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(3)...", Label1(4)
    
    cSQL = "Insert into Artikel Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", AGN "
    cSQL = cSQL & ", PGN "
    cSQL = cSQL & ", LEKNEU as LEKPR "
    cSQL = cSQL & ", VKPR "
    cSQL = cSQL & ", MWST "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", EAN2"
    cSQL = cSQL & ", EAN3 "
    cSQL = cSQL & ", ETIMERK "
'    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", 0 as BESTAND "
    cSQL = cSQL & ", MINMEN "
    cSQL = cSQL & ", INHALT "
    cSQL = cSQL & ", INHALTBEZ "
    cSQL = cSQL & ", GRUNDPREIS "
    cSQL = cSQL & ", MINBEST "
    cSQL = cSQL & ", 'J' as RABATT_OK "
    cSQL = cSQL & ", 'J' as GEFUEHRT "
    cSQL = cSQL & ", EKPR "
    cSQL = cSQL & ", 'N' as PREISSCHU "
    cSQL = cSQL & ", 'J' as BONUS_OK "
    cSQL = cSQL & ", 'J' as UMS_OK "
    cSQL = cSQL & ", AWM "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as LASTTIME "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as AUFDAT "
    cSQL = cSQL & ", EXDAT "
    cSQL = cSQL & ", GROESSE "
    cSQL = cSQL & ", 'A' as SYNSTATUS "
    cSQL = cSQL & ", NOTIZEN "
    cSQL = cSQL & ", KVKNEU as KVKPR1 "
    cSQL = cSQL & " from ImportpriREWE where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(4)...", Label1(4)
    
    cSQL = "Insert into ARTLIEF Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LEKNEU as LEKPR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", MINMEN "
    cSQL = cSQL & ", 0 as SPANNE "
    cSQL = cSQL & ", 'A' as SYNSTATUS "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & " from ImportpriREWE where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(5)...", Label1(4)
    
    
    'Neuheiten
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROREWE Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'Neuheiten/wieder verf¸gbar' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
    cSQL = cSQL & ", KVKNEU as VKPR_NEW "
    cSQL = cSQL & " from ImportpriREWE where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(6)...", Label1(4)
    
    'EK-Preis‰nderungen
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROREWE Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'EK-Preis‰nderungen' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
    cSQL = cSQL & ", LEKALT as LEK_ALT "
    cSQL = cSQL & ", LEKNEU as LEK_NEW "
    cSQL = cSQL & " from ImportpriREWE where LEKALT <> 0 and RKZ = 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(7)...", Label1(4)
    
'    'alle anderen ƒnderungen + Artliefeintrag
'    cSQL = "Update Artikel a inner join ImportpriREWE i on a.artnr = i.artnr "
'    cSQL = cSQL & " set a.EAN = i.EAN  "
'    cSQL = cSQL & " , a.EAN2 = a.EAN "
'    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
'    cSQL = cSQL & " where i.EAN <> a.ean "
'    cSQL = cSQL & " and not a.ean is null"
'    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    'Wir haben 2 EAN zur Verf¸gung
    
    
    
    'Anfang neu
    
    'noch weiter hinten
    
    
''    test fang von hinten an
'    '3.EAN
'    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8.1)...", Label1(4)
'    loeschNEW "REWEZWEAN", gdBase
'    CreateTableT2 "REWEZWEAN", gdBase
'
'    cSQL = "Insert into REWEZWEAN select ARTNR,val(EAN3) as EANTO from ImportpriREWE "
'    cSQL = cSQL & " where (not ean3 is null or ean3 = '')"
'    gdBase.Execute cSQL, dbFailOnError
'
'    cSQL = "Update REWEZWEAN set EANTO = '0' & EANTO  "
'    cSQL = cSQL & " where len(EANTO)= 11 "
'    gdBase.Execute cSQL, dbFailOnError
'
'    cSQL = "Delete from REWEZWEAN  "
'    cSQL = cSQL & " where EANTO = '0' "
'    gdBase.Execute cSQL, dbFailOnError
'
'    cSQL = "Update Artikel a inner join REWEZWEAN i on a.artnr = i.artnr "
'    cSQL = cSQL & " set a.EAN = i.EANTO  "
'    cSQL = cSQL & " , a.EAN2 = a.EAN "
'    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
'    cSQL = cSQL & " where i.EANTO <> a.ean "
'    cSQL = cSQL & " and not a.ean is null"
'    gdBase.Execute cSQL, dbFailOnError











    'fang alle vorhandenen EAN auf
    loeschNEW "REWEUrEAN", gdBase
    cSQL = "Create Table REWEUrEAN (ARTNR int, EANUR Text(13))"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into REWEUrEAN select a.ARTNR,val(a.EAN) as EANUR from ImportpriREWE i inner join Artikel a on i.artnr = a.artnr"
    cSQL = cSQL & " where (not a.ean is null or a.ean = '')"
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Insert into REWEUrEAN select a.ARTNR,val(a.EAN2) as EANUR from ImportpriREWE i inner join Artikel a on i.artnr = a.artnr"
    cSQL = cSQL & " where (not a.ean2 is null or a.ean2 = '')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into REWEUrEAN select a.ARTNR,val(a.EAN3) as EANUR from ImportpriREWE i inner join Artikel a on i.artnr = a.artnr"
    cSQL = cSQL & " where (not a.ean3 is null or a.ean3 = '')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into REWEUrEAN select a.ARTNR,val(a.EAN) as EANUR from ImportpriREWE i inner join artean_k a on i.artnr = a.artnr"
    cSQL = cSQL & " where (not a.ean is null or a.ean = '')"
    gdBase.Execute cSQL, dbFailOnError


    cSQL = "Delete from REWEUrEAN where eanur = '0'"
    gdBase.Execute cSQL, dbFailOnError






    
    '2.EAN
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8.2)...", Label1(4)
    loeschNEW "REWEZWEAN", gdBase
    CreateTableT2 "REWEZWEAN", gdBase
    
    cSQL = "Insert into REWEZWEAN select ARTNR,val(EAN2) as EANTO from ImportpriREWE "
    cSQL = cSQL & " where (not ean2 is null or ean2 = '')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REWEZWEAN set EANTO = '0' & EANTO  "
    cSQL = cSQL & " where len(EANTO)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from REWEZWEAN  "
    cSQL = cSQL & " where EANTO = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel a inner join REWEZWEAN i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN = i.EANTO  "
    cSQL = cSQL & " , a.EAN2 = a.EAN "
    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
    cSQL = cSQL & " where i.EANTO <> a.ean "
    cSQL = cSQL & " and not a.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    '1.EAN
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8.3)...", Label1(4)
    loeschNEW "REWEZWEAN", gdBase
    CreateTableT2 "REWEZWEAN", gdBase
    
    cSQL = "Insert into REWEZWEAN select ARTNR,val(EAN) as EANTO from ImportpriREWE "
    cSQL = cSQL & " where (not ean is null or ean = '')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REWEZWEAN set EANTO = '0' & EANTO  "
    cSQL = cSQL & " where len(EANTO)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from REWEZWEAN  "
    cSQL = cSQL & " where EANTO = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel a inner join REWEZWEAN i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN = i.EANTO  "
    cSQL = cSQL & " , a.EAN2 = a.EAN "
    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
    cSQL = cSQL & " where i.EANTO <> a.ean "
    cSQL = cSQL & " and not a.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    
    'ende neu
    
    
    'Jetzt alles lˆschen aus Reweurean, was bekannt ist
    
'    cSQL = "Create Table REWEUrEAN (ARTNR int, EANUR Text(13))"
'    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = " Alter table REWEUrEAN add erkannt Text(1)  "
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Update REWEUrEAN set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REWEUrEAN inner join  ARTIKEL on REWEUrEAN.EANUR =  ARTIKEL.ean and REWEUrEAN.artnr =  ARTIKEL.artnr "
    cSQL = cSQL & " set REWEUrEAN.erkannt = 'J'"
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Delete from REWEUrEAN where erkannt = 'J'  "
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Update REWEUrEAN set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update REWEUrEAN inner join  ARTIKEL on REWEUrEAN.EANUR =  ARTIKEL.ean2 and REWEUrEAN.artnr =  ARTIKEL.artnr "
    cSQL = cSQL & " set REWEUrEAN.erkannt = 'J'"
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Delete from REWEUrEAN where erkannt = 'J'  "
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Update REWEUrEAN set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REWEUrEAN inner join  ARTIKEL on REWEUrEAN.EANUR =  ARTIKEL.ean3 and REWEUrEAN.artnr =  ARTIKEL.artnr "
    cSQL = cSQL & " set REWEUrEAN.erkannt = 'J'"
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Delete from REWEUrEAN where erkannt = 'J'  "
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Update REWEUrEAN set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REWEUrEAN inner join  Artean_K on REWEUrEAN.EANUR =  Artean_K.ean and REWEUrEAN.artnr =  Artean_K.artnr "
    cSQL = cSQL & " set REWEUrEAN.erkannt = 'J'"
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Delete from REWEUrEAN where erkannt = 'J'  "
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Update REWEUrEAN set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    'jetzt versuch auf ean2 abzusetzen
    


    Dim rsArt As DAO.Recordset
    Dim cEAN As String
    Dim cArtNr As String
    Dim bEingef¸gt As Boolean
    Dim cEANART As String
    
    cSQL = "Select distinct(eanur), Artnr from REWEUrEAN"
    
    Set rsRS = gdBase.OpenRecordset(cSQL)
    If Not rsRS.EOF Then
        
        
        rsRS.MoveFirst
        Do While Not rsRS.EOF
        
            bEingef¸gt = False
            

            If Not IsNull(rsRS!eanur) Then
                cEAN = Trim(rsRS!eanur)
            End If
            
            If Not IsNull(rsRS!artnr) Then
                cArtNr = Trim(rsRS!artnr)
            End If

            cSQL = "Select * from Artikel where artnr = " & cArtNr & ""
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                
                rsArt.Edit
                
                If Not IsNull(rsArt!EAN) Then
                    cEANART = Trim(rsArt!EAN)
                    If Val(cEANART) = 0 Then
                        rsArt!EAN = cEAN
                        bEingef¸gt = True
                    End If
                Else
                    rsArt!EAN = cEAN
                    bEingef¸gt = True
                End If
                
                If bEingef¸gt = False Then
                
                    If Not IsNull(rsArt!EAN2) Then
                        cEANART = Trim(rsArt!EAN2)
                        If Val(cEANART) = 0 Then
                            rsArt!EAN2 = cEAN
                            bEingef¸gt = True
                        End If
                    Else
                        rsArt!EAN2 = cEAN
                        bEingef¸gt = True
                    End If
                
                End If
                
                If bEingef¸gt = False Then
                
                    If Not IsNull(rsArt!EAN3) Then
                        cEANART = Trim(rsArt!EAN3)
                        If Val(cEANART) = 0 Then
                            rsArt!EAN3 = cEAN
                            bEingef¸gt = True
                        End If
                    Else
                        rsArt!EAN3 = cEAN
                        bEingef¸gt = True
                    End If
                
                End If
                
                If bEingef¸gt = False Then
                
                    cSQL = "Insert into Artean_k (artnr,ean) values (" & cArtNr & ",'" & cEAN & "')"
                    gdBase.Execute cSQL, dbFailOnError
                End If
                
                

                
                rsArt.Update
                
            End If
            rsArt.Close: Set rsArt = Nothing
            rsRS.MoveNext
        Loop
    End If

    rsRS.Close: Set rsRS = Nothing
    
    
    
    
        
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8.4)...", Label1(4)
    
    cSQL = "Update Artikel a inner join ImportpriREWE i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.BEZEICH = i.BEZEICH "
    cSQL = cSQL & " , a.LIBESNR = i.LIBESNR "
'    cSQL = cSQL & " , a.EAN = i.EAN "
    cSQL = cSQL & " , a.MWST = i.MWST "
    cSQL = cSQL & " , a.MINMEN = i.MINMEN "
    cSQL = cSQL & " , a.INHALT = i.INHALT "
    cSQL = cSQL & " , a.INHALTBEZ = i.INHALTBEZ "
    cSQL = cSQL & " , a.AGN = i.AGN "
    cSQL = cSQL & " , a.NOTIZEN = i.NOTIZEN "
    cSQL = cSQL & " , a.lekpr = i.lekneu"
    cSQL = cSQL & ", a.LASTDATE = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", a.LASTTIME = '" & TimeValue(Now) & "' "
    cSQL = cSQL & ", a.SYNSTATUS = 'E' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(9)...", Label1(4)
    
    cSQL = "Delete from artlief where artnr in (Select artnr from importprirewe) and Linr = 400001 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into ARTLIEF Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LEKNEU as LEKPR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", MINMEN "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & ", NULL as EXDAT "
    cSQL = cSQL & ", 0 as SPANNE "
    cSQL = cSQL & ", 'E' as SYNSTATUS "
    cSQL = cSQL & " from ImportpriREWE " 'where awm <> '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    'Anfang neu
    cSQL = "Update Artikel set EAN = '' where EAN = '0' "
    cSQL = cSQL & " and artnr in (Select artnr from ImportpriREWE) "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update Artikel set EAN2 = '' where EAN2 = '0' "
    cSQL = cSQL & " and artnr in (Select artnr from ImportpriREWE) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN3 = '' where EAN3 = '0' "
    cSQL = cSQL & " and artnr in (Select artnr from ImportpriREWE) "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    

    cSQL = "update artikel set ean3 = '' where ean3 = ean "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update artikel set ean2 = '' where ean2 = ean "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update artikel set ean3 = '' where ean3 = ean2 "
    gdBase.Execute cSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("artean_K", gdBase) Then

        If SpalteInTabellegefundenNEW("artean_K", "erkannt", gdBase) = False Then
            cSQL = " Alter table artean_K add erkannt Text(1)  "
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        
        
        'gibt es den EAN aus Artean_K auch an EAN1 der Artikel
    
        cSQL = "Update artean_K set erkannt = 'N'  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean and artean_K.artnr =  ARTIKEL.artnr "
        cSQL = cSQL & " set artean_K.erkannt = 'J'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from artean_K where erkannt = 'J'  "
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        cSQL = "Update artean_K set erkannt = 'N'  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean2 and artean_K.artnr =  ARTIKEL.artnr "
        cSQL = cSQL & " set artean_K.erkannt = 'J'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from artean_K where erkannt = 'J'  "
        gdBase.Execute cSQL, dbFailOnError
        
        
        cSQL = "Update artean_K set erkannt = 'N'  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean3 and artean_K.artnr =  ARTIKEL.artnr "
        cSQL = cSQL & " set artean_K.erkannt = 'J'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from artean_K where erkannt = 'J'  "
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        
        
        
    
'        cSQL = "Update artean_K set erkannt = 'N'  "
'        gdBase.Execute cSQL, dbFailOnError
'
'        cSQL = "Update artean_K inner join  ARTIKEL on artean_K.artnr =  ARTIKEL.artnr "
'        cSQL = cSQL & " set artean_K.erkannt = 'J'"
'        gdBase.Execute cSQL, dbFailOnError
'
'        cSQL = "Delete from artean_K where erkannt = 'N'  "
'        gdBase.Execute cSQL, dbFailOnError
        
        If SpalteInTabellegefundenNEW("artean_K", "erkannt", gdBase) = True Then
            cSQL = " Alter table artean_K drop erkannt   "
            gdBase.Execute cSQL, dbFailOnError
        End If
    
    End If
    'ende neu
    
    
    
    cSQL = "Update Artlief inner join ImportpriREWE i on Artlief.artnr = i.artnr  "
    cSQL = cSQL & " set Artlief.RKZ = 'J'"
    cSQL = cSQL & ", Artlief.EXDAT = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", Artlief.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    cSQL = cSQL & "  and Artlief.linr = 400001 "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(10)...", Label1(4)
    
    cSQL = "Update STADAPROREWE s inner join Artikel a  on s.artnr = a.artnr "
    cSQL = cSQL & " set s.farbnr = val(a.awm) "
    cSQL = cSQL & " , s.agn = a.agn "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(11)...", Label1(4)
    
    cSQL = "Update STADAPROREWE s inner join AGNDBF a  on s.agn = a.agn "
    cSQL = cSQL & " set s.agtext = a.agtext "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(12)...", Label1(4)
    
    'Hier VK-Preisanpassungen
    cSQL = "Update Artikel a inner join ImportpriREWE i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.kvkpr1 = i.vkpr  "
    cSQL = cSQL & " where i.mnotizen = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(13)...", Label1(4)
    
    BringFarbeInsSpiel "STADAPROREWE", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebernahme_Rewe_Delta"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Sicherheitslˆschen_mitLinr(sArtnr As String, sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Delete from artlief where artnr = " & sArtnr & " and Linr = " & sLinr
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Sicherheitslˆschen_mitLinr"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Status_Ermitteln()
On Error GoTo LOKAL_ERROR

    Dim rs As Recordset
    Dim sSQL As String
    
    'neue Artikel
    sSQL = "Select count(*) as maxi from ImportpriREWE where awm = '98'"
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!maxi) Then
            Label1(6).Caption = CInt(Label1(6).Caption) + Val(Trim(rs!maxi))
            If Val(Label1(6).Caption) > 0 Then
                Command5(6).Visible = True
                Command5(6).Enabled = False
            End If
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
    'EK Preis‰nderungen
    sSQL = "Select count(*) as maxi from ImportpriREWE where LEKALT <> 0"
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!maxi) Then
            Label1(8).Caption = CInt(Label1(8).Caption) + Val(Trim(rs!maxi))
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
   
    'VK Preis‰nderungen
    sSQL = "Select count(*) as maxi from ImportpriREWE where Round(KVKALT,2) <> Round(KVKNEU,2)"
    sSQL = sSQL & " and mnotizen = 'J' "
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!maxi) Then
            Label1(13).Caption = CInt(Label1(13).Caption) + Val(Trim(rs!maxi))
            If Val(Label1(13).Caption) > 0 Then
                Command5(5).Visible = True
                Command5(5).Enabled = False
            End If
        End If
        
    End If
    rs.Close: Set rs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "status_ermitteln"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErmittlungReweDuplisPlusDel()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsRS        As Recordset
    Dim rsArt       As Recordset
    Dim cEAN        As String
    Dim lcount      As Long
    
    lcount = 0
    
    loeschNEW "ImportDupli", gdApp
    
    sSQL = "select count(ean) as count ,ean into ImportDupli from Meister group by ean having count(ean) > 1"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  ImportDupli where ean is null"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  ImportDupli where trim(ean) = ''"
    gdApp.Execute sSQL, dbFailOnError
    
    
    Set rsRS = gdApp.OpenRecordset("ImportDupli", dbOpenTable)
    If Not rsRS.EOF Then
        
        
        rsRS.MoveFirst
        Do While Not rsRS.EOF
        
            

            If Not IsNull(rsRS!EAN) Then
                cEAN = Trim(rsRS!EAN)
            End If

            sSQL = "Select * from Meister where ean = '" & cEAN & "'"
            Set rsArt = gdApp.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst

                rsArt.MoveNext
                Do While Not rsArt.EOF

                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsRS.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If

    rsRS.Close: Set rsRS = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ErmittlungReweDuplisPlusDel"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub feldcheckRewe()
    On Error GoTo LOKAL_ERROR
    
    Dim rsMeister       As Recordset
    Dim rsIMPORTPRI     As Recordset
    Dim sSQL            As String

    Dim sBez            As String
    Dim sEAN            As String
    
    
    loeschNEW "ImportpriREWE", gdApp
    CreateTableT2 "IMPORTPRIREWE", gdApp
    
    Set rsIMPORTPRI = gdApp.OpenRecordset("ImportpriREWE")
    
    Set rsMeister = gdApp.OpenRecordset("Meister")
    If Not rsMeister.EOF Then
    
        
        rsMeister.MoveFirst
        Do While Not rsMeister.EOF
        
            rsIMPORTPRI.AddNew
            rsIMPORTPRI!artnr = rsMeister!artnr
            
            'Bezeichnung auf * und ' checken
            If Not IsNull(rsMeister!BEZEICH) Then
                sBez = Left(rsMeister!BEZEICH, 35)
                sBez = SwapStr(sBez, "*", " ")  'stern
                sBez = SwapStr(sBez, Chr(34), "Z") '"
                sBez = SwapStr(sBez, "'", " ")  'Hochkommata
                sBez = SwapStr(sBez, ",", ".")  'komma
                sBez = SwapStr(sBez, "·", "ﬂ")  'ﬂ
                sBez = SwapStr(sBez, "î", "ˆ")  '
                sBez = SwapStr(sBez, "Ñ", "‰")  '
                sBez = SwapStr(sBez, "Å", "¸")
                sBez = SwapStr(sBez, "ô", "÷")
                sBez = SwapStr(sBez, "ö", "‹")
                sBez = SwapStr(sBez, "`", " ")  '
                sBez = SwapStr(sBez, "|", " ")  '
            Else
                sBez = ""
            End If
            If gbTagAkt = True Then
                rsIMPORTPRI!BEZEICH = UCase(sBez)
            Else
                rsIMPORTPRI!BEZEICH = sBez
            End If
            'Standardm‰ﬂig auf Ziffer 0
            
            If Not IsNull(rsMeister!AGN) Then
                rsIMPORTPRI!AGN = rsMeister!AGN
            Else
                rsIMPORTPRI!AGN = 0
            End If
            
            If Not IsNull(rsMeister!PGN) Then
                rsIMPORTPRI!PGN = rsMeister!PGN
            Else
                rsIMPORTPRI!PGN = 0
            End If
            
            If Not IsNull(rsMeister!RKZ) Then
                rsIMPORTPRI!RKZ = rsMeister!RKZ
            Else
                rsIMPORTPRI!RKZ = "N"
            End If
            
            rsIMPORTPRI!lekneu = Round(rsMeister!lekpr, 2)
            rsIMPORTPRI!vkpr = rsMeister!vkpr
            
            If rsMeister!MWST = "1" Then
                rsIMPORTPRI!MWST = "V"
            ElseIf rsMeister!MWST = "2" Then
                rsIMPORTPRI!MWST = "E"
            ElseIf rsMeister!MWST = "" Then
                rsIMPORTPRI!MWST = "V"
            ElseIf IsNull(rsMeister!MWST) Then
                rsIMPORTPRI!MWST = "V"
            Else
                rsIMPORTPRI!MWST = rsMeister!MWST
            End If
            
            rsIMPORTPRI!linr = rsMeister!linr
            rsIMPORTPRI!LIBESNR = rsMeister!LIBESNR
            
            If Not IsNull(rsMeister!EAN) Then
                sEAN = rsMeister!EAN
            Else
                sEAN = "0"
            End If
            
            sEAN = SwapStr(sEAN, "00000", "")
            
            If sEAN = "" Then
                rsIMPORTPRI!EAN = "0"
            Else
                rsIMPORTPRI!EAN = sEAN
            End If

            rsIMPORTPRI!EAN2 = rsMeister!EAN2
            rsIMPORTPRI!EAN3 = rsMeister!EAN3
            rsIMPORTPRI!ETIMERK = rsMeister!ETIMERK
            rsIMPORTPRI!MOPREIS = rsMeister!MOPREIS
            
            'standardm‰ﬂig auf "N"

            rsIMPORTPRI!NOTIZEN = rsMeister!NOTIZEN
            rsIMPORTPRI!BESTAND = rsMeister!BESTAND
            rsIMPORTPRI!VKMENGE = rsMeister!VKMENGE
            rsIMPORTPRI!VKDATUM = rsMeister!VKDATUM
            rsIMPORTPRI!MINMEN = rsMeister!MINMEN
            rsIMPORTPRI!INHALT = rsMeister!INHALT
            rsIMPORTPRI!INHALTBEZ = Trim(UCase(rsMeister!INHALTBEZ))
            rsIMPORTPRI!GRUNDPREIS = rsMeister!GRUNDPREIS
            rsIMPORTPRI!MINBEST = rsMeister!MINBEST
            rsIMPORTPRI!RABATT_OK = rsMeister!RABATT_OK
            rsIMPORTPRI!GEFUEHRT = rsMeister!GEFUEHRT
            rsIMPORTPRI!KVKPR1 = rsMeister!KVKPR1
            rsIMPORTPRI!ekpr = rsMeister!ekpr
            rsIMPORTPRI!PREISSCHU = rsMeister!PREISSCHU
            rsIMPORTPRI!BONUS_OK = rsMeister!BONUS_OK
            rsIMPORTPRI!UMS_OK = rsMeister!UMS_OK
            rsIMPORTPRI!AWM = rsMeister!AWM
            rsIMPORTPRI!LASTDATE = rsMeister!LASTDATE
            rsIMPORTPRI!LASTTIME = rsMeister!LASTTIME
            rsIMPORTPRI!MNOTIZEN = rsMeister!Status
            rsIMPORTPRI!KVKalt = 0
            rsIMPORTPRI!KVKNEU = rsMeister!vkpr
            
            If Not IsNull(rsMeister!AUFDAT) Then
                rsIMPORTPRI!AUFDAT = DateValue(Right(rsMeister!AUFDAT, 2) & "." & Mid(rsMeister!AUFDAT, 5, 2) & "." & Left(rsMeister!AUFDAT, 4))
            Else
                rsIMPORTPRI!AUFDAT = Null
            End If
            
            If Not IsNull(rsMeister!EXDAT) Then
                rsIMPORTPRI!EXDAT = DateValue(Right(rsMeister!EXDAT, 2) & "." & Mid(rsMeister!EXDAT, 5, 2) & "." & Left(rsMeister!EXDAT, 4))
            Else
                rsIMPORTPRI!EXDAT = Null
            End If
            
            rsIMPORTPRI!LPZ = rsMeister!LPZ
            rsIMPORTPRI!FARBNR = rsMeister!FARBNR
            rsIMPORTPRI!MARKE = rsMeister!MARKE
            rsIMPORTPRI!GROESSE = rsMeister!GROESSE
            rsIMPORTPRI!SPANNE = rsMeister!SPANNE
            rsIMPORTPRI!AUFSCHLAG = rsMeister!AUFSCHLAG
            rsIMPORTPRI!SYNStatus = rsMeister!SYNStatus
           
            rsIMPORTPRI.Update
            rsMeister.MoveNext
        Loop
    End If
    rsMeister.Close
    rsIMPORTPRI.Close
    
    sSQL = "Delete from ImportpriREWE where EAN is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ImportpriREWE where bezeich is null"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ImportpriREWE where bezeich = ''"
    gdApp.Execute sSQL, dbFailOnError
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "feldcheckRewe"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Function FormatiereBildschirmdatenREWE() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsRS            As Recordset
    Dim lreservArtnr    As Long
    Dim lvergebeArtnr   As Long
   
    FormatiereBildschirmdatenREWE = False
    
    anzeige "normal", "Neue Artikel werden ermittelt...", Label1(4)
    'Farbe alle auf neu
    sSQL = "Update ImportpriREWE set AWM = '98' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    anzeige "normal", "Neue Artikel werden ermittelt(1)......", Label1(4)
    
    sSQL = "Create Index AWM on ImportpriREWE (AWM)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(2).........", Label1(4)
    
    sSQL = "Create Index ean on ImportpriREWE (ean)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(3)............", Label1(4)
    
    sSQL = "Create Index linr on ImportpriREWE (linr)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(4)...", Label1(4)
    
    sSQL = "Create Index libesnr on ImportpriREWE (libesnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(5)......", Label1(4)
    
    
    'Artikel mit EAN ¸bereinstimmung auf standard
    sSQL = "Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(6).........", Label1(4)
    
    
    
    sSQL = "Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN2 = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(7)............", Label1(4)
    
    
    
    sSQL = "Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN3 = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(8)...", Label1(4)
    
    
    sSQL = "Update ImportpriREWE i inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = i.LINR and ARTLIEF.LIBESNR = i.LIBESNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.artnr = ARTLIEF.artnr  "
    sSQL = sSQL & " where i.artnr is null "
    sSQL = sSQL & " and i.LIBESNR <> '0000000' "
    gdBase.Execute sSQL, dbFailOnError

    anzeige "normal", "Neue Artikel werden ermittelt(9)......", Label1(4)

    

    sSQL = "Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = i.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  = '0' and i.AWM = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(10).........", Label1(4)
    
    'Lekpreisver‰nderungen anzeigen
    
    
    sSQL = "Update ImportpriREWE i inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = i.LINR and ARTLIEF.artnr = i.artnr "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.LEKALT = ARTLIEF.LEKPR "
    sSQL = sSQL & " where not i.artnr is null "
    sSQL = sSQL & " and i.rkz  = 'N' "
'    sSQL = sSQL & " and Round(i.lekneu,2) = Round(ARTLIEF.LEKPR,2) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select LEKALT from ImportpriREWE where not LEKALT is null "
    Set rsRS = gdBase.OpenRecordset(sSQL)
    If Not rsRS.EOF Then
        rsRS.MoveFirst
        Do While Not rsRS.EOF
            rsRS.Edit
            rsRS!LEKALT = Round(rsRS!LEKALT, 2)
            rsRS.Update
            
            rsRS.MoveNext
        Loop
    
    End If
    rsRS.Close: Set rsRS = Nothing
    
    sSQL = "Update ImportpriREWE set LEKALT = 0 "
    sSQL = sSQL & " where not artnr is null "
    sSQL = sSQL & " and rkz  = 'N' "
    sSQL = sSQL & " and Round(lekneu,2) = Round(LEKALT,2) "
    gdBase.Execute sSQL, dbFailOnError
    
    'Ende Lekpreisver‰nderungen anzeigen
    
   
    
    anzeige "normal", "Neue Artikel werden ermittelt(12)...", Label1(4)
    
    sSQL = " Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = i.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.PREISSCHU = ARTIKEL.PREISSCHU "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(13)......", Label1(4)
    
    
    
    Status_Ermitteln
    
    
    'Zieh doch mal  den freien Artikelnummernkreis hoch
    
    
    anzeige "normal", "F¸r neue Artikel werden freie Artikelnummern ermittelt...", Label1(4)
    
    lreservArtnr = HoleFreieArtikelNrab(glartv, glartb)
    
    
    sSQL = "Select * from ImportpriREWE where awm = '98' "
    Set rsRS = gdBase.OpenRecordset(sSQL)
    If Not rsRS.EOF Then
        rsRS.MoveFirst
        Do While Not rsRS.EOF
            rsRS.Edit
            rsRS!artnr = lreservArtnr
            rsRS.Update
            
            lvergebeArtnr = NextfreieArtnr(lreservArtnr, glartb)
            If lvergebeArtnr = 0 Then
                anzeige "rot", "Es stehen keine neuen Artikelnummern zur Verf¸gung (Einstellungen ¸berpr¸fen).", Label1(4)
                Exit Function
            Else
                lreservArtnr = lvergebeArtnr
                
            End If
            rsRS.MoveNext
        Loop
    
    End If
    rsRS.Close: Set rsRS = Nothing
    
    If lvergebeArtnr > 0 Then
        sSQL = "Update FFE set ARTNRV = " & lvergebeArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    
    FormatiereBildschirmdatenREWE = True
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereBildschirmdatenREWE"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function
Private Sub Command5_Click(Index As Integer)
 On Error GoTo LOKAL_ERROR
 
    Dim sdatname    As String
    Dim i           As Integer
    Dim sSQL        As String
    Dim lLfnr       As Long
    Dim cLfnr       As String
    Dim rsRS        As DAO.Recordset
 
    Select Case Index
    
        Case 0
            voreinstellungspeichernE185
            Unload frmWKL185
        Case 1      'Rewe Stammdaten einlesen
        
            Timer1.Enabled = False
            Command5(1).Enabled = False
            
            Dim cSQL As String
            If NewTableSuchenDBKombi("STADAPROREWE", gdBase) = True Then
                cSQL = "Delete from STADAPROREWE "
                gdBase.Execute cSQL, dbFailOnError
            End If
            
            'Ablaufprotokoll f¸llen
            'Etiketten erstellen
            'dem Anwender ein ‹bernahmeergebnis zeigen
            
            CreateTableT2 "RORDER", gdBase
            
            File1.Path = gsKinPfad 'Standard In Pfad
            File1.Pattern = "REWE*Delta.csv"
            File1.Refresh
            
            If File1.ListCount > 0 Then
                'Datei/en stehen an
                For i = 0 To File1.ListCount - 1
                    sdatname = File1.list(i)
                    cLfnr = Mid(sdatname, 10, 3)
                    lLfnr = Val(cLfnr)
                    
                    sSQL = "Insert into RORDER (lfnr,DATNAME)"
                    sSQL = sSQL & " Values ( "
                    sSQL = sSQL & " " & lLfnr & " "
                    sSQL = sSQL & ", '" & sdatname & "' "
                    sSQL = sSQL & " ) "
                    gdBase.Execute sSQL, dbFailOnError
                Next i
            End If
            
            sSQL = "Select * from RORDER order by lfnr asc"
            Set rsRS = gdBase.OpenRecordset(sSQL)
            If Not rsRS.EOF Then
                rsRS.MoveFirst
                Do While Not rsRS.EOF
                    If Not IsNull(rsRS!Datname) Then
                        sdatname = rsRS!Datname
                        ReweDatenEinlesen gsKinPfad, sdatname
                    End If
                
                rsRS.MoveNext
                Loop
            End If
            rsRS.Close: Set rsRS = Nothing
            
            loeschNEW "RORDER", gdBase
            
        Case 2
            anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
            reportbildschirm "", "aWKL185"
            
            
        Case 3
        
            'EX Artikel als Etiketten
            
            Screen.MousePointer = 11
            
            loeschNEW "LSTEETI", gdBase
            CreateTableT2 "LSTEETI", gdBase
            
            sSQL = "Insert into LSTEETI select Artnr "
            sSQL = sSQL & ", BEZEICH "
            sSQL = sSQL & ", 1 as BESTAND "
            sSQL = sSQL & ", 1 as ANZAHL "
            sSQL = sSQL & ", KVKNEU as VKPR "
            sSQL = sSQL & ", LIBESNR "
            sSQL = sSQL & ", EAN "
            sSQL = sSQL & ", LPZ "
            sSQL = sSQL & ", LINR "
            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
            sSQL = sSQL & " from ImportpriREWE_Dauer "
            sSQL = sSQL & " Where rkz = 'J'"
            gdBase.Execute sSQL, dbFailOnError

            gbEtiExArtikel = True

            gsETILS = "aus Lieferschein"
            
        
            frmWKL30.Show 1
            
            gbEtiExArtikel = False
            
        Case 5
        
            'VK Pflichtpreisanpassungen als Etiketten
            
            Screen.MousePointer = 11
            
            loeschNEW "LSTEETI", gdBase
            CreateTableT2 "LSTEETI", gdBase
            
            sSQL = "Insert into LSTEETI select Artnr "
            sSQL = sSQL & ", BEZEICH "
            sSQL = sSQL & ", 1 as BESTAND "
            sSQL = sSQL & ", 1 as ANZAHL "
            sSQL = sSQL & ", KVKNEU as VKPR "
            sSQL = sSQL & ", LIBESNR "
            sSQL = sSQL & ", EAN "
            sSQL = sSQL & ", LPZ "
            sSQL = sSQL & ", LINR "
            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
            sSQL = sSQL & " from ImportpriREWE_Dauer "
            sSQL = sSQL & " Where Round(KVKALT, 2) <> Round(KVKNEU, 2)"
            sSQL = sSQL & " and mnotizen = 'J' "
            gdBase.Execute sSQL, dbFailOnError

            gsETILS = "aus Lieferschein"
        
            frmWKL30.Show 1
            
        Case 6
        
            'neue Artikel als Etiketten
            
            Screen.MousePointer = 11
            
            loeschNEW "LSTEETI", gdBase
            CreateTableT2 "LSTEETI", gdBase
            
            sSQL = "Insert into LSTEETI select Artnr "
            sSQL = sSQL & ", BEZEICH "
            sSQL = sSQL & ", 1 as BESTAND "
            sSQL = sSQL & ", 1 as ANZAHL "
            sSQL = sSQL & ", KVKNEU as VKPR "
            sSQL = sSQL & ", LIBESNR "
            sSQL = sSQL & ", EAN "
            sSQL = sSQL & ", LPZ "
            sSQL = sSQL & ", LINR "
            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
            sSQL = sSQL & " from ImportpriREWE_Dauer "
            sSQL = sSQL & " where awm = '98' "
            gdBase.Execute sSQL, dbFailOnError

            gsETILS = "aus Lieferschein"
        
            frmWKL30.Show 1
        
    
    End Select
    
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    
    If NewTableSuchenDBKombi("E185", gdBase) Then
        
        voreinstellungladenE185
    
    End If
    
    
    lesenEinstellungen
    iSec = 0
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE185()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdBase.OpenRecordset("E185")
    If Not rs.EOF Then
    
        If rs!bo1 = True Then
            chk_Pflichtpreisanpassungen_vornehmen.Value = vbUnchecked
        Else
            chk_Pflichtpreisanpassungen_vornehmen.Value = vbChecked
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE185"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE185()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim bo1     As Integer
    
    loeschNEW "E185", gdBase
    CreateTableT2 "E185", gdBase
    
    If chk_Pflichtpreisanpassungen_vornehmen.Value = vbChecked Then
        bo1 = 0
    Else
        bo1 = -1
    End If
    
    sSQL = "Insert into E185 ( bo1) "
    sSQL = sSQL & " values (" & bo1 & ")"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE185"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lesenEinstellungen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsRS        As Recordset
    
    glartv = 600000
    glartb = 700000
    
    If NewTableSuchenDBKombi("FFE", gdBase) = True Then
        Set rsRS = gdBase.OpenRecordset("FFE", dbOpenTable)
        If Not rsRS.EOF Then
            rsRS.MoveFirst
            
            If Not IsNull(rsRS!ARTNRV) Then
                glartv = rsRS!ARTNRV
            Else
                glartv = 600000
            End If
            
            If Not IsNull(rsRS!ARTNRB) Then
                glartb = rsRS!ARTNRB
            Else
                glartb = 700000
            End If
        End If
        rsRS.Close: Set rsRS = Nothing
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lesenEinstellungen"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "REWE_AGN", gdApp
    loeschNEW "MEISTER", gdApp
'    loeschNEW "STADAPROREWE", gdBase
    loeschNEW "ImportDupli", gdApp
    loeschNEW "ImportpriREWE", gdBase
    loeschNEW "ImportpriREWE_dauer", gdBase
    loeschNEW "ImportpriREWE", gdApp
    
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
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Timer1.Enabled = False

If Index = 11 Then
    URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/hilfe-bei-problemen/44-software-probleme-winkiss/231-rewe-stammdaten-einlesen.html"
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer1_Timer()
On Error GoTo LOKAL_ERROR

    iSec = iSec + 1
    
    If iSec >= 10 Then
        Unload frmWKL185
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil neue REWE Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
