VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL195 
   BackColor       =   &H00C0C000&
   Caption         =   "neue L¸ning - Artikeldaten"
   ClientHeight    =   8610
   ClientLeft      =   1215
   ClientTop       =   1590
   ClientWidth     =   11910
   Icon            =   "frmWKL195.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chk_vk 
      Caption         =   "nur relevante Artikel, die schon einmal verkauft wurden"
      Height          =   495
      Left            =   8400
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   10560
      Pattern         =   "MASTER!.*"
      TabIndex        =   6
      Top             =   1920
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
      MouseIcon       =   "frmWKL195.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   15
      ToolTipText     =   "hier alle Neuigkeiten lesen"
      Top             =   5880
      Visible         =   0   'False
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
      Caption         =   "Neue L¸ning-Artikeldaten stehen bereit. "
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
      Caption         =   "neue L¸ning - Artikeldaten"
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
Attribute VB_Name = "frmWKL195"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim glartv      As Long
Dim glartb      As Long
Dim iSec        As Integer
Dim lAgNforZig  As Long
Private Sub L¸ningDatenEinlesen(sPfad As String, sDatei As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim rsEAN           As DAO.Recordset
    Dim rsEX            As DAO.Recordset
    Dim lAnz            As Long
    Dim lLenfil         As Long
    Dim lPosEnde        As Long
    Dim lposSemi        As Long
    Dim lposSemiEnde    As Long
    Dim cWert           As String
    Dim cInhalt         As String
    Dim cInhaltBez      As String
    Dim i               As Integer
    Dim iFileNr         As Integer
    Dim cSatz1          As String
    Dim lPos            As Long
    Dim cEinzelsatz     As String
    Dim lLinr           As Long
    Dim lLiefnr         As Long
    Dim lfnr1           As Long
    Dim ctemp           As String
    Dim lagn            As Long
    Dim sRKZ            As String
        
    lfnr1 = 0
    lPos = 1
    
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
    
    If Not NewTableSuchenDBKombi("LUEE", gdBase) Then 'das erste Mal
        CreateTableT2 "LUEE", gdBase
    End If
    
    lLinr = checkL¸ninginLISRT()
    
    Screen.MousePointer = 11
    
    loeschapp "meister"
    CreateTable "MEISTER", gdApp

    loeschNEW "EAN61", gdApp
    CreateTableT2 "EAN61", gdApp
    
    loeschNEW "EX82", gdApp
    CreateTableT2 "EX82", gdApp
    
    Set rsEX = gdApp.OpenRecordset("EX82")
    Set rsEAN = gdApp.OpenRecordset("EAN61")
    Set rsrs = gdApp.OpenRecordset("MEISTER")

    iFileNr = FreeFile
    Open sPfad & "\" & sDatei For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
    
        cSatz1 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz1
    
        lLenfil = Len(cSatz1)
        
        lPosEnde = InStr(lPos, cSatz1, vbCrLf)

        Do
            lPosEnde = InStr(lPos, cSatz1, vbCrLf)
            cEinzelsatz = Mid(cSatz1, lPos, lPosEnde - lPos)
            lPos = lPos + lPosEnde - lPos + 2
            
            If Left(cEinzelsatz, 2) = "81" Then
                rsrs.AddNew
                rsrs!LIBESNR = Val(Mid(cEinzelsatz, 7, 8))
                rsrs!vkpr = Mid(cEinzelsatz, 66, 7)
                rsrs!lekpr = Mid(cEinzelsatz, 141, 6)
                rsrs!BEZEICH = Mid(cEinzelsatz, 84, 30)

                ctemp = ""
                ctemp = Trim(Mid(cEinzelsatz, 334, 10))
                'den numerischen Teil abschneiden
                cInhaltBez = ""
                cInhalt = ""
                If Len(ctemp) > 0 Then
                    For i = Len(ctemp) To 0 Step -1
                        cInhalt = Mid(ctemp, 1, i)
                        If IsNumeric(cInhalt) Then
                            cInhaltBez = Right(ctemp, Len(ctemp) - i)
                            cInhaltBez = SwapStr(cInhaltBez, ".", "")
                            Exit For
                        End If
                    Next i
                End If
                
                rsrs!INHALTBEZ = UCase(Left(cInhaltBez, 3))
                
                If cInhalt <> "" Then
                    rsrs!INHALT = cInhalt
                End If
                
                If cInhalt <> "" And cInhaltBez <> "" Then
                    rsrs!GRUNDPREIS = "J"
                Else
                    rsrs!GRUNDPREIS = "N"
                End If
                
                rsrs!MINMEN = Mid(cEinzelsatz, 73, 4)
                
                ctemp = ""
                ctemp = Mid(cEinzelsatz, 33, 1)
                If ctemp = "2" Then
                    rsrs!MWST = "V"
                ElseIf ctemp = "1" Then
                    rsrs!MWST = "E"
                ElseIf ctemp = "0" Then
                    rsrs!MWST = "O"
                End If
                
                rsrs!LPZ = 1
                rsrs!GEFUEHRT = "J"
                rsrs!EAN = ""
                rsrs!EAN2 = ""
                rsrs!EAN3 = ""
                rsrs!linr = lLinr
                
                lagn = Val(Mid(cEinzelsatz, 54, 4))
                
                
                If giLuening = -1 Then 'ohne Lebensmittel
                    If (lagn >= 800 And lagn < 1299) Or (lagn >= 1500 And lagn < 9799) Or lagn = 431 Then '431 Corny-Riegel
                        lfnr1 = lfnr1 + 1
                        rsrs!lfnr = lfnr1
                        
                        
                        
                        If lagn >= 1250 And lagn < 1260 Then
                            rsrs!AGN = lAgNforZig
                        Else
                            rsrs!AGN = 617
                        End If
                        
                        
                        
                        
'                        rsrs!AGN = 617
                        
                        ctemp = ""
                        ctemp = Trim(Mid(cEinzelsatz, 21, 1))
                        If ctemp = "A" Then
                            rsrs!RKZ = "N"
                        Else
                            rsrs!RKZ = "J"
                            Label1(7).Caption = CInt(Label1(7).Caption) + 1
                            Label1(7).Refresh
                        End If
                        
                        rsrs.Update
                    End If
                ElseIf giLuening = 0 Then 'alles auch Lebensmittel
                    lfnr1 = lfnr1 + 1
                    rsrs!lfnr = lfnr1
                    
                    If lagn >= 1250 And lagn < 1260 Then
                        rsrs!AGN = lAgNforZig
                    Else
                        rsrs!AGN = 617
                    End If
                    
'                    rsrs!AGN = 617
                    
                    ctemp = ""
                    ctemp = Trim(Mid(cEinzelsatz, 21, 1))
                    If ctemp = "A" Then
                        rsrs!RKZ = "N"
                    Else
                        rsrs!RKZ = "J"
                        Label1(7).Caption = CInt(Label1(7).Caption) + 1
                        Label1(7).Refresh
                    End If
                    
                    rsrs.Update
                End If

            ElseIf Left(cEinzelsatz, 2) = "61" Then
            
                rsEAN.AddNew
                rsEAN!LIBESNR = Val(Mid(cEinzelsatz, 7, 8))
                rsEAN!EAN = Trim(Mid(cEinzelsatz, 33, 13) & fn_errechne_Pr¸fziffer(Trim(Mid(cEinzelsatz, 33, 13))))
                
                rsEAN.Update
                
            ElseIf Left(cEinzelsatz, 2) = "82" Then
            
                rsEX.AddNew
                
                
                ctemp = ""
                ctemp = Trim(Mid(cEinzelsatz, 21, 1))
                If ctemp = "L" Then
                    rsEX.AddNew
                    rsEX!LIBESNR = Val(Mid(cEinzelsatz, 7, 8))
                    rsEX.Update
                ElseIf ctemp = "S" Then
                    rsEX.AddNew
                    rsEX!LIBESNR = Val(Mid(cEinzelsatz, 7, 8))
                    rsEX.Update
                ElseIf ctemp = "R" Then
                    rsEX.AddNew
                    rsEX!LIBESNR = Val(Mid(cEinzelsatz, 7, 8))
                    rsEX.Update
                End If
                
                
            ElseIf Left(cEinzelsatz, 2) = "54" Then
            
                ctemp = Val(Mid(cEinzelsatz, 34, 6))
                If Val(ctemp) > 0 Then
            
                    rsrs.AddNew
                    
                    lfnr1 = lfnr1 + 1
                    rsrs!lfnr = lfnr1
                    rsrs!LIBESNR = Val(Mid(cEinzelsatz, 3, 12))
                    rsrs!BEZEICH = "54"
                    rsrs!EAN = lfnr1
                    rsrs!lekpr = ctemp
                    rsrs!linr = lLinr
                    
                    rsrs.Update
                    
                End If
            
            End If
            
        Loop While lLenfil >= lPos
    End If
    
    Close iFileNr
    rsrs.Close: Set rsrs = Nothing
    rsEAN.Close: Set rsEAN = Nothing
    
    
    
    
    loeschNEW "artean", gdApp
    
    sSQL = "Create Index libesnr on meister (libesnr)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index libesnr on EAN61 (libesnr)"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from EAN61 where right(ean,1) = 'x' "
    gdApp.Execute sSQL, dbFailOnError
    
    
    'hier zwischen ARTEAN_K f¸llen
    loeschNEW "artean_LUE", gdApp
    
    sSQL = "Select EAN61.libesnr, EAN61.ean, 0 as Artnr into artean_LUE from EAN61"
    gdApp.Execute sSQL, dbFailOnError
    
    loesch "artean_LUE"
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "artean_LUE"
    'ende hier zwischen ARTEAN_K f¸llen
    
    'hier zwischen EX82 f¸llen
    
    loeschNEW "EX82", gdBase
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "EX82"
    'ende hier zwischen EX82 f¸llen
    
    
    
    
    
    
    

    
    
    sSQL = "Select EAN61.libesnr, EAN61.ean into artean from EAN61 where EAN61.libesnr in (Select meister.libesnr from meister)"
    gdApp.Execute sSQL, dbFailOnError
    
    loeschNEW "Top1ean", gdApp

    sSQL = "Select max(Ean) as eanMax, libesnr into Top1ean from artean group by libesnr"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update meister inner join top1ean on meister.libesnr = top1ean.libesnr set meister.ean = top1ean.eanMax "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from artean where ean in (Select ean from Top1ean)"
    gdApp.Execute sSQL, dbFailOnError
    
    loeschNEW "Top1ean", gdApp

    sSQL = "Select max(Ean) as eanMax, libesnr into Top1ean from artean group by libesnr"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update meister inner join top1ean on meister.libesnr = top1ean.libesnr set meister.ean2 = top1ean.eanMax "
    gdApp.Execute sSQL, dbFailOnError
    
    
     anzeige "normal", "EAN Duplikate werden entfernt...", Label1(4)
    '6. EAN - Duplikats¸berpr¸fung in der Importtabelle Anzahl ermitteln
    ErmittlungLueningDuplisPlusDel
    
    anzeige "normal", "Die Datens‰tze werden ¸berpr¸ft...", Label1(4)
    '7. diverse Feld¸berpr¸fungen vornehmen
    feldcheckLuening
    anzeige "normal", "Die Datens‰tze werden verarbeitet...", Label1(4)

    '8. Tabelle IMPORTPRI zur Datenbank kopieren
    loesch "ImportpriLuening"
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "IMPORTPRILuening"

    sSQL = "Create Index ARTNR on IMPORTPRILuening (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError
    
    
    '54
    
    'vorgreifen
    sSQL = "Update IMPORTPRILuening inner join Artlief on "
    sSQL = sSQL & " IMPORTPRILuening.libesnr = Artlief.libesnr "
    sSQL = sSQL & "Set IMPORTPRILuening.artnr = Artlief.artnr  "
    sSQL = sSQL & ", IMPORTPRILuening.MINMEN = Artlief.MINMEN  "
    
    sSQL = sSQL & ", IMPORTPRILuening.LEKALT = Round(Artlief.LEKPR,2)  "
    
    sSQL = sSQL & " where Artlief.linr = " & lLinr
    sSQL = sSQL & " and IMPORTPRILuening.bezeich = '54' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from  IMPORTPRILuening where bezeich = '54' and artnr is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update IMPORTPRILuening set bezeich = '', ean = '' where bezeich = '54' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update IMPORTPRILuening inner join Artikel on "
    sSQL = sSQL & " IMPORTPRILuening.Artnr = Artikel.artnr "
    sSQL = sSQL & "Set IMPORTPRILuening.Bezeich = Artikel.Bezeich  "
    sSQL = sSQL & ", IMPORTPRILuening.EAN = Artikel.EAN  "
    sSQL = sSQL & ", IMPORTPRILuening.AGN = Artikel.AGN  "
    sSQL = sSQL & ", IMPORTPRILuening.PGN = Artikel.PGN  "
    sSQL = sSQL & ", IMPORTPRILuening.GRUNDPREIS = Artikel.GRUNDPREIS  "
    sSQL = sSQL & ", IMPORTPRILuening.INHALT = Artikel.INHALT  "
    sSQL = sSQL & ", IMPORTPRILuening.INHALTBEZ = Artikel.INHALTBEZ  "
    sSQL = sSQL & ", IMPORTPRILuening.MWST = Artikel.MWST  "
    
    sSQL = sSQL & ", IMPORTPRILuening.gefuehrt = 'J' "
    sSQL = sSQL & ", IMPORTPRILuening.RKZ = 'N' "
    sSQL = sSQL & ", IMPORTPRILuening.LPZ = 1 "
    sSQL = sSQL & ", IMPORTPRILuening.AWM = '0'"
    sSQL = sSQL & ", IMPORTPRILuening.KVKALT = Artikel.KVKPR1  "
    sSQL = sSQL & ", IMPORTPRILuening.KVKNEU = Artikel.KVKPR1  "
    sSQL = sSQL & ", IMPORTPRILuening.vkpr = Artikel.KVKPR1  "
    
    sSQL = sSQL & " where not IMPORTPRILuening.artnr is null"
    gdBase.Execute sSQL, dbFailOnError
    
 
    
    '54 Ende
    
    FormatiereBildschirmdatenLuening

     'jetzt ¸bernehmen
    Uebernahme_Luening_Delta sDatei, lLinr
    
    If NewTableSuchenDBKombi("ImportpriLuening_dauer", gdBase) = False Then
        sSQL = "select * into ImportpriLuening_dauer from ImportpriLuening  "
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Insert into ImportpriLuening_dauer select * from ImportpriLuening "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Kill sPfad & "\" & sDatei
    
    Command5(2).Enabled = True
    Command5(3).Enabled = True
    Command5(5).Enabled = True
    Command5(6).Enabled = True
    
    anzeige "normal", "Fertig! Artikel¸bernahme beendet, Protokoll und Etiketten ausdrucken!", Label1(4)
    
    Screen.MousePointer = 0
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "L¸ningDatenEinlesen"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Private Function FormatiereBildschirmdatenLuening() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim lreservArtnr    As Long
    Dim lvergebeArtnr   As Long
    
    FormatiereBildschirmdatenLuening = False
    
    anzeige "normal", "Neue Artikel werden ermittelt...", Label1(4)
    'Farbe alle auf neu
    sSQL = "Update IMPORTPRILuening set AWM = '98' "
    sSQL = sSQL & " where IMPORTPRILuening.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(1)......", Label1(4)
    
    sSQL = "Create Index AWM on IMPORTPRILuening (AWM)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(2).........", Label1(4)
    
    sSQL = "Create Index ean on IMPORTPRILuening (ean)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(3)............", Label1(4)
    
    sSQL = "Create Index linr on IMPORTPRILuening (linr)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(4)...", Label1(4)
    
    sSQL = "Create Index libesnr on IMPORTPRILuening (libesnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(5)......", Label1(4)
    
    'Artikel mit EAN ¸bereinstimmung auf standard
    sSQL = "Update IMPORTPRILuening inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN = IMPORTPRILuening.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRILuening.AWM = '0' "
    sSQL = sSQL & " ,IMPORTPRILuening.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,IMPORTPRILuening.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,IMPORTPRILuening.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,IMPORTPRILuening.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,IMPORTPRILuening.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where IMPORTPRILuening.ean  <> '0' "
    sSQL = sSQL & " and IMPORTPRILuening.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(6).........", Label1(4)
    
    sSQL = "Update IMPORTPRILuening inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN2 = IMPORTPRILuening.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRILuening.AWM = '0' "
    sSQL = sSQL & " ,IMPORTPRILuening.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,IMPORTPRILuening.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,IMPORTPRILuening.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,IMPORTPRILuening.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,IMPORTPRILuening.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where IMPORTPRILuening.ean  <> '0' "
    sSQL = sSQL & " and IMPORTPRILuening.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(7)............", Label1(4)
    
    sSQL = "Update IMPORTPRILuening inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN3 = IMPORTPRILuening.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRILuening.AWM = '0' "
    sSQL = sSQL & " ,IMPORTPRILuening.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,IMPORTPRILuening.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,IMPORTPRILuening.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,IMPORTPRILuening.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,IMPORTPRILuening.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where IMPORTPRILuening.ean  <> '0' "
    sSQL = sSQL & " and IMPORTPRILuening.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Update IMPORTPRILuening inner join ARTEAN_K on "
    sSQL = sSQL & " ARTEAN_K.EAN = IMPORTPRILuening.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRILuening.AWM = '0' "
    sSQL = sSQL & " ,IMPORTPRILuening.artnr = ARTEAN_K.artnr  "
    sSQL = sSQL & " where IMPORTPRILuening.ean  <> '0' "
    sSQL = sSQL & " and IMPORTPRILuening.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    anzeige "normal", "Neue Artikel werden ermittelt(8)...", Label1(4)
    
    sSQL = "Update IMPORTPRILuening inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = IMPORTPRILuening.LINR and ARTLIEF.LIBESNR = IMPORTPRILuening.LIBESNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRILuening.AWM = '0' "
    sSQL = sSQL & " ,IMPORTPRILuening.artnr = ARTLIEF.artnr  "
    sSQL = sSQL & " where IMPORTPRILuening.artnr is null "
'    sSQL = sSQL & " where Importpri.ean  = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    sSQL = "Delete from IMPORTPRILuening "
    sSQL = sSQL & " where IMPORTPRILuening.artnr is null "
    sSQL = sSQL & " and  IMPORTPRILuening.ean is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from IMPORTPRILuening "
    sSQL = sSQL & " where IMPORTPRILuening.artnr is null "
    sSQL = sSQL & " and  IMPORTPRILuening.ean = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from IMPORTPRILuening "
    sSQL = sSQL & " where IMPORTPRILuening.artnr is null "
    sSQL = sSQL & " and  IMPORTPRILuening.ean = '0' "
    gdBase.Execute sSQL, dbFailOnError

    anzeige "normal", "Neue Artikel werden ermittelt(9)......", Label1(4)

    sSQL = "Update IMPORTPRILuening inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = IMPORTPRILuening.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRILuening.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,IMPORTPRILuening.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,IMPORTPRILuening.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where IMPORTPRILuening.ean  = '0' and IMPORTPRILuening.AWM = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(10).........", Label1(4)
    
    'Lekpreisver‰nderungen anzeigen
    sSQL = "Update IMPORTPRILuening i inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = i.LINR and ARTLIEF.artnr = i.artnr "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.LEKALT = ARTLIEF.LEKPR "
    sSQL = sSQL & " where not i.artnr is null "
    sSQL = sSQL & " and i.rkz  = 'N' "
'    sSQL = sSQL & " and Round(i.lekneu,2) = Round(ARTLIEF.LEKPR,2) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select LEKALT from IMPORTPRILuening where not LEKALT is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!LEKALT = Round(rsrs!LEKALT, 2)
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Update IMPORTPRILuening set LEKALT = 0 "
    sSQL = sSQL & " where not artnr is null "
    sSQL = sSQL & " and rkz  = 'N' "
    sSQL = sSQL & " and Round(lekneu,2) = Round(LEKALT,2) "
    gdBase.Execute sSQL, dbFailOnError
    'Ende Lekpreisver‰nderungen anzeigen
    
    sSQL = " Update IMPORTPRILuening inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = IMPORTPRILuening.ARTNR "
    sSQL = sSQL & " Set "
    sSQL = sSQL & " IMPORTPRILuening.PREISSCHU = ARTIKEL.PREISSCHU "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(13)......", Label1(4)
    
    sSQL = "Delete from IMPORTPRILuening where bezeich is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from IMPORTPRILuening where bezeich = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    Status_Ermitteln
    
    anzeige "normal", "F¸r neue Artikel werden freie Artikelnummern ermittelt...", Label1(4)
    
    lreservArtnr = HoleFreieArtikelNrab(glartv, glartb)

    sSQL = "Select * from IMPORTPRILuening where awm = '98' or awm = '95' or awm = '94'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!artnr = lreservArtnr
            rsrs.Update
            
            lvergebeArtnr = NextfreieArtnr(lreservArtnr, glartb)
            If lvergebeArtnr = 0 Then
                anzeige "rot", "Es stehen keine neuen Artikelnummern zur Verf¸gung (Einstellungen ¸berpr¸fen).", Label1(4)
                Exit Function
            Else
                lreservArtnr = lvergebeArtnr
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    FormatiereBildschirmdatenL¸ning = True
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereBildschirmdatenLuening"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function
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
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Status_Ermitteln()
On Error GoTo LOKAL_ERROR

    Dim rs As Recordset
    Dim sSQL As String
    
    'neue Artikel
    sSQL = "Select count(*) as maxi from ImportpriLuening where awm = '98'"
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
    sSQL = "Select count(*) as maxi from ImportpriLuening where LEKALT <> 0"
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            Label1(8).Caption = CInt(Label1(8).Caption) + Val(Trim(rs!maxi))
        End If
    End If
    rs.Close: Set rs = Nothing
    
    'VK Preis‰nderungen
    sSQL = "Select count(*) as maxi from ImportpriLuening where Round(KVKALT,2) <> Round(KVKNEU,2)"
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
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErmittlungLueningDuplisPlusDel()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
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
    
    Set rsrs = gdApp.OpenRecordset("ImportDupli", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!EAN) Then
                cEAN = Trim(rsrs!EAN)
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
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ErmittlungLueningDuplisPlusDel"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
 On Error GoTo LOKAL_ERROR
 
    Dim sdatname    As String
    Dim i           As Integer
    Dim sSQL        As String
    Dim lLfnr       As Long
    Dim cLfnr       As String
    Dim rsrs        As DAO.Recordset
 
    Select Case Index
    
        Case 0
            Unload frmWKL195
        Case 1      'L¸ning Stammdaten einlesen
        
            Timer1.Enabled = False
            Command5(1).Enabled = False
        
            'Ablaufprotokoll f¸llen
            'Etiketten erstellen
            'dem Anwender ein ‹bernahmeergebnis zeigen
            
            If NewTableSuchenDBKombi("STADAPROLUENING", gdBase) = True Then
                cSQL = "Delete from STADAPROLUENING "
                gdBase.Execute cSQL, dbFailOnError
            End If
            
            CreateTableT2 "LORDER", gdBase
            
            File1.Path = gsKinPfad 'Standard In Pfad
            File1.Pattern = "A00*.dat"
            File1.Refresh
            
            If File1.ListCount > 0 Then
                'Datei/en stehen an
                For i = 0 To File1.ListCount - 1
                    sdatname = File1.list(i)
                    cLfnr = Mid(sdatname, 8, 8)
                    lLfnr = Val(cLfnr)
                    
                    sSQL = "Insert into LORDER (lfnr,DATNAME)"
                    sSQL = sSQL & " Values ( "
                    sSQL = sSQL & " " & lLfnr & " "
                    sSQL = sSQL & ", '" & sdatname & "' "
                    sSQL = sSQL & " ) "
                    gdBase.Execute sSQL, dbFailOnError
                Next i
            End If
            
            sSQL = "Select * from LORDER order by lfnr asc"
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                Do While Not rsrs.EOF
                    If Not IsNull(rsrs!Datname) Then
                        sdatname = rsrs!Datname
                        
                        L¸ningDatenEinlesen gsKinPfad, sdatname
                        
                    End If
                rsrs.MoveNext
                Loop
            End If
            rsrs.Close: Set rsrs = Nothing
            
            loeschNEW "LORDER", gdBase
            
        Case 2
            anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
            reportbildschirm "", "aWKL195"
            
        Case 3  'EX Artikel als Etiketten
        
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
            sSQL = sSQL & " from ImportpriLuening_Dauer "
            sSQL = sSQL & " Where rkz = 'J'"
            
            If chk_vk.Value = vbChecked Then
                sSQL = sSQL & " and ImportpriLuening_Dauer.Artnr in (Select Artnr from Kassjour)"
            End If
            
            gdBase.Execute sSQL, dbFailOnError

            gbEtiExArtikel = True

            gsETILS = "aus Lieferschein"
        
            frmWKL30.Show 1
            
            gbEtiExArtikel = False
            
        Case 5  'VK Pflichtpreisanpassungen als Etiketten
            
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
            sSQL = sSQL & " from ImportpriLuening_Dauer "
            sSQL = sSQL & " Where Round(KVKALT, 2) <> Round(KVKNEU, 2)"
            sSQL = sSQL & " and mnotizen = 'J' "
            gdBase.Execute sSQL, dbFailOnError

            gsETILS = "aus Lieferschein"
            frmWKL30.Show 1
            
        Case 6  'neue Artikel als Etiketten
        
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
            sSQL = sSQL & " from ImportpriLuening_Dauer "
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
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub feldcheckLuening()
    On Error GoTo LOKAL_ERROR
    
    Dim rsMeister       As Recordset
    Dim rsIMPORTPRI     As Recordset
    Dim sSQL            As String
    Dim sBez            As String
    Dim sEAN            As String
    
    loeschNEW "ImportpriLuening", gdApp
    CreateTableT2 "IMPORTPRILUENING", gdApp
    
    Set rsIMPORTPRI = gdApp.OpenRecordset("ImportpriLuening")
    
    Set rsMeister = gdApp.OpenRecordset("Meister")
    If Not rsMeister.EOF Then
        
        rsMeister.MoveFirst
        Do While Not rsMeister.EOF
        
            rsIMPORTPRI.AddNew
            rsIMPORTPRI!lfnr = rsMeister!lfnr
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
            
            rsIMPORTPRI!lekneu = rsMeister!lekpr / 100
            rsIMPORTPRI!vkpr = rsMeister!vkpr / 100
            
            
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
            rsIMPORTPRI!KVKNEU = rsMeister!vkpr / 100
            
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
    
    sSQL = "Delete from ImportpriLuening where EAN is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ImportpriLuening where bezeich is null"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ImportpriLuening where bezeich = ''"
    gdApp.Execute sSQL, dbFailOnError
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "feldcheckLuening"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    checkFFE
    lesenEinstellungenFFE
    iSec = 0
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function checkL¸ninginLISRT() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim rsLi As Recordset
    
    checkL¸ninginLISRT = 0

    sSQL = "Select linr from LUEE "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!linr) Then
            checkL¸ninginLISRT = rsrs!linr
            
            sSQL = "Select * from LISRT where LINR = " & checkL¸ninginLISRT
            sSQL = sSQL & " and ( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )"
            Set rsLi = gdBase.OpenRecordset(sSQL)
            If rsLi.RecordCount = 0 Then
                checkL¸ninginLISRT = 0
            End If
            rsLi.Close
        
        End If
    End If
    
    If checkL¸ninginLISRT = 0 Then
    
        Screen.MousePointer = 0
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "LINR"
        If gF2Prompt.cFeld <> "" Then
            gsAnzeige00a = "Bitte den L¸ning - Lieferant ausw‰hlen!"
            frmWK00a.Show 1
        End If
        gsAnzeige00a = ""
        
        anzeige "normal", "Der Lieferant: " & gF2Prompt.cWahl & " wurde zugeordnet.", Label1(4)
        
        If gF2Prompt.cWahl <> "" Then
             checkL¸ninginLISRT = CDbl(gF2Prompt.cWahl)
        End If
        
        If checkL¸ninginLISRT <> 0 Then
            sSQL = "update LUEE set linr = " & checkL¸ninginLISRT
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkL¸ninginLISRT"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub lesenEinstellungenFFE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    glartv = 600000
    glartb = 700000
    lAgNforZig = 0
    
    
    If NewTableSuchenDBKombi("FFE", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("FFE", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!ARTNRV) Then
                glartv = rsrs!ARTNRV
            Else
                glartv = 600000
            End If
            
            If Not IsNull(rsrs!ARTNRB) Then
                glartb = rsrs!ARTNRB
            Else
                glartb = 700000
            End If
            
            If Not IsNull(rsrs!AGNLUE) Then
                lAgNforZig = rsrs!AGNLUE
            Else
                lAgNforZig = 0
            End If
            
            If Not IsNull(rsrs!LUENING) Then
                If rsrs!LUENING = True Then
                    giLuening = 0
                Else
                    giLuening = -1
                End If
            Else
                giLuening = -1
            End If
            
        End If
        rsrs.Close: Set rsrs = Nothing
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lesenEinstellungenFFE"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "MEISTER", gdApp
'    loeschNEW "STADAPROLUENING", gdBase
    loeschNEW "ImportDupli", gdApp
    loeschNEW "ImportpriLuening", gdBase
    loeschNEW "ImportpriLuening_dauer", gdBase
    loeschNEW "ImportpriLuening", gdApp
    
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
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer1_Timer()
On Error GoTo LOKAL_ERROR

    iSec = iSec + 1
    
    If iSec >= 10 Then
        Unload frmWKL195
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Uebernahme_Luening_Delta(sDatei As String, lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim rsArtlief       As DAO.Recordset
    Dim sArtnr          As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen...", Label1(4)
    
    If NewTableSuchenDBKombi("STADAPROLUENING", gdBase) = False Then
        CreateTableT2 "STADAPROLUENING", gdBase
    End If
    
    
    
    'Alle EANS auffangen************************************************

    loeschNEW "artean_BU", gdBase

    cSQL = "Create Table artean_BU  "
    cSQL = cSQL & "( ARTNR int"
    cSQL = cSQL & ", EANCH varchar(13)"
    cSQL = cSQL & ", OTHERARTNR1 int"
    cSQL = cSQL & ", OTHERARTNR1_Bestand int"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError


    cSQL = "Insert into artean_BU Select artnr, val(ean) as EANCH from ImportpriLuening  "
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
    cSQL = "Update Artikel inner join ImportpriLuening i on artikel.artnr = i.artnr "
    cSQL = cSQL & " set "
    cSQL = cSQL & " artikel.LASTDATE = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", artikel.LASTTIME = '" & TimeValue(Now) & "' "
    cSQL = cSQL & ", artikel.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artlief inner join ImportpriLuening i on Artlief.artnr = i.artnr  "
    cSQL = cSQL & " set Artlief.RKZ = 'J'"
    cSQL = cSQL & ", Artlief.EXDAT = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", Artlief.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    cSQL = cSQL & "  and Artlief.linr = " & lLinr
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(1)...", Label1(4)
    
    If Val(Label1(7).Caption) > 0 Then
        Command5(3).Visible = True
        Command5(3).Enabled = False
        chk_vk.Visible = True
    End If
    
    'Ex Artikel
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROLUENING Select "
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
    cSQL = cSQL & " from ImportpriLuening where rkz = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(2)...", Label1(4)
    
    'neue Artikel
    cSQL = "Select * from ImportpriLuening where awm = '98'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!artnr) Then
                sArtnr = Trim(rsrs!artnr)
            End If
            
            Sicherheitslˆschen sArtnr 'artlief
            rsrs.MoveNext
            
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
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
    cSQL = cSQL & " from ImportpriLuening where awm = '98' "
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
    cSQL = cSQL & " from ImportpriLuening where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(5)...", Label1(4)
    
    
    'Neuheiten
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROLUENING Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'Neuheiten/wieder verf¸gbar' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
    cSQL = cSQL & ", KVKNEU as VKPR_NEW "
    cSQL = cSQL & " from ImportpriLuening where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(6)...", Label1(4)
    
    'EK-Preis‰nderungen
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROLUENING Select "
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
    cSQL = cSQL & " from ImportpriLuening where LEKALT <> 0 and RKZ = 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(7)...", Label1(4)
    
    'alle anderen ƒnderungen + Artliefeintrag
    cSQL = "Update Artikel a inner join ImportpriLuening i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN = i.EAN  "
    cSQL = cSQL & " , a.EAN2 = a.EAN "
    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
    cSQL = cSQL & " where i.EAN <> a.ean "
    cSQL = cSQL & " and not a.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(7.1)...", Label1(4)
    

    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8)...", Label1(4)
    
    cSQL = "Update Artikel a inner join ImportpriLuening i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.BEZEICH = i.BEZEICH "
    cSQL = cSQL & ", a.LIBESNR = i.LIBESNR "
    cSQL = cSQL & ", a.EAN = i.EAN "
    cSQL = cSQL & ", a.MWST = i.MWST "
    cSQL = cSQL & ", a.MINMEN = i.MINMEN "
    cSQL = cSQL & ", a.INHALT = i.INHALT "
    cSQL = cSQL & ", a.INHALTBEZ = i.INHALTBEZ "
'    cSQL = cSQL & ", a.AGN = i.AGN "
    cSQL = cSQL & ", a.NOTIZEN = i.NOTIZEN "
    cSQL = cSQL & ", a.lekpr = i.lekneu"
    cSQL = cSQL & ", a.LASTDATE = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", a.LASTTIME = '" & TimeValue(Now) & "' "
    cSQL = cSQL & ", a.SYNSTATUS = 'E' "
    gdBase.Execute cSQL, dbFailOnError
    
    'Hier Artikel reaktivieren

    'Erst Protokoll f¸llen
    cSQL = "Insert into STADAPROLUENING Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'Neuheiten/wieder verf¸gbar' as AKT "
    cSQL = cSQL & ", i.ARTNR "
    cSQL = cSQL & ", i.BEZEICH "
    cSQL = cSQL & ", i.EAN "
    cSQL = cSQL & ", i.libesnr "
    cSQL = cSQL & ", i.KVKNEU as VKPR_NEW "
    cSQL = cSQL & " from ImportpriLuening i inner join Artikel a on i.artnr = a.artnr "
    cSQL = cSQL & " where i.rkz = 'N' and a.rkz = 'J'"
    gdBase.Execute cSQL, dbFailOnError

    'Dann Inhalt ¸bernehmen
    cSQL = "Update ARTLIEF a inner join ImportpriLuening i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.RKZ = i.RKZ "
    cSQL = cSQL & ", a.SYNSTATUS = 'E' "
    cSQL = cSQL & ", a.Exdat = 0 "
    cSQL = cSQL & " where i.rkz = 'N' and a.rkz = 'J' and a.Linr = " & lLinr
    gdBase.Execute cSQL, dbFailOnError


    'Hier Artikel reaktivieren _ Ende***********************
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(9)...", Label1(4)
    
    cSQL = "Delete from artlief where artnr in (Select artnr from ImportpriLuening) and Linr = " & lLinr
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
    cSQL = cSQL & " from ImportpriLuening " 'where awm <> '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artlief inner join ImportpriLuening i on Artlief.artnr = i.artnr  "
    cSQL = cSQL & " set Artlief.RKZ = 'J'"
    cSQL = cSQL & ", Artlief.EXDAT = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", Artlief.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    cSQL = cSQL & "  and Artlief.linr = " & lLinr
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(10)...", Label1(4)
    
    cSQL = "Update STADAPROLUENING s inner join Artikel a  on s.artnr = a.artnr "
    cSQL = cSQL & " set s.farbnr = val(a.awm) "
    cSQL = cSQL & " , s.agn = a.agn "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(11)...", Label1(4)
    
    cSQL = "Update STADAPROLUENING s inner join AGNDBF a  on s.agn = a.agn "
    cSQL = cSQL & " set s.agtext = a.agtext "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(12)...", Label1(4)
    
'    'Hier VK-Preisanpassungen
'    cSQL = "Update Artikel a inner join ImportpriREWE i on a.artnr = i.artnr "
'    cSQL = cSQL & " set a.kvkpr1 = i.vkpr  "
'    cSQL = cSQL & " where i.mnotizen = 'J' "
'    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Update Artikel a inner join ImportpriLuening i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN = '' where a.EAN = '0'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel a inner join ImportpriLuening i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN2 = '' where a.EAN2 = '0'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel a inner join ImportpriLuening i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN3 = '' where a.EAN3 = '0'"
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(13)...", Label1(4)
    
    BringFarbeInsSpiel "STADAPROLUENING", gdBase
    
    
     'hier die vielen EAN beim L¸ning Einleseverfahren in die ARTEAN_K speichern
    
'    artean_LUE
    If NewTableSuchenDBKombi("artean_LUE", gdBase) = True Then
    
        cSQL = "Update artean_LUE a inner join ImportpriLuening i on a.libesnr = i.Libesnr "
        cSQL = cSQL & " set a.artnr = i.artnr  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update artean_LUE set artnr = 0 where artnr is null"
        gdBase.Execute cSQL, dbFailOnError
    
        cSQL = "Delete from artean_LUE where artnr = 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        If Not NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then 'das erste Mal
    
            cSQL = "Create Table ARTEAN_K"
            cSQL = cSQL & " ( "
            cSQL = cSQL & " ARTNR long "
            cSQL = cSQL & ", ean Text(13) "
            cSQL = cSQL & " ) "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Create Index EAN on ARTEAN_K (EAN)"
            gdBase.Execute cSQL, dbFailOnError
            
        End If
        
        
        'Artnr + EAN Kombinationen
        
        'Alle EANS auffangen************************************************

        loeschNEW "artean_BU", gdBase
    
        cSQL = "Create Table artean_BU  "
        cSQL = cSQL & "( ARTNR int"
        cSQL = cSQL & ", EANCH varchar(13)"
        cSQL = cSQL & ", OTHERARTNR1 int"
        cSQL = cSQL & ", OTHERARTNR1_Bestand int"
        cSQL = cSQL & ") "
        gdBase.Execute cSQL, dbFailOnError
    
        cSQL = "Insert into artean_BU Select artnr, val(ean) as EANCH from artean_LUE  "
        gdBase.Execute cSQL, dbFailOnError
    
        cSQL = "Update artean_BU set EANCH = '0' & EANCH  "
        cSQL = cSQL & " where len(EANCH)= 11 "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update artean_BU set OTHERARTNR1 = 0 , OTHERARTNR1_Bestand = 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from artean_BU where EANCH = '0' "
        gdBase.Execute cSQL, dbFailOnError
        
        If NewTableSuchenDBKombi("artean_Artikel", gdBase) Then
    

    
    
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
        
        
        End If
        
        
    
        'Ende ********************** Alle EANS auffangen
        
        
        
        
        'Ende Ende
        
        cSQL = "Insert into ARTEAN_K select artnr, ean from  artean_LUE where artnr > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "artean_LUE", gdBase
    
    End If
    
    
    '    artean_LUE
    If NewTableSuchenDBKombi("EX82", gdBase) = True Then
    
        cSQL = "Update Artlief inner join EX82 i on Artlief.libesnr = i.libesnr  "
        cSQL = cSQL & " set Artlief.RKZ = 'J'"
        cSQL = cSQL & ", Artlief.EXDAT = '" & DateValue(Now) & "' "
        cSQL = cSQL & ", Artlief.SYNSTATUS = 'E' "
        cSQL = cSQL & " where Artlief.linr = " & lLinr
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If lAgNforZig > 0 Then
        cSQL = "Update Artikel set rabatt_OK = 'N' where agn = " & lAgNforZig
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebernahme_Luening_Delta"
    Fehler.gsFehlertext = "Im Programmteil neue L¸ning Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

