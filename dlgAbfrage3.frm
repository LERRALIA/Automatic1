VERSION 5.00
Begin VB.Form dlgAbfrage3 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   2400
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5895
   Icon            =   "dlgAbfrage3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Pfeil
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin sevCommand3.Command cmdWiederholen 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Wiederholen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin sevCommand3.Command cmdAusführen 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Überschreiben"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin sevCommand3.Command CancelButton 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Abbrechen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblBeschriftung 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "dlgAbfrage3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_Beschriftung As String
Private m_Buttoneinscaption As String
Private m_Buttonzweicaption As String
Private m_Buttondreicaption As String
Private m_Rückgabewert As Integer
Private m_Überschrift As String


Public Property Let Beschriftung(Schrift As String)
    m_Beschriftung = Schrift
End Property

Public Property Get Beschriftung() As String
    Beschriftung = m_Beschriftung
End Property

Public Property Let BCaptioneins(Schrift As String)
    m_Buttoneinscaption = Schrift
End Property

Public Property Get BCaptioneins() As String
    BCaptioneins = m_Buttoneinscaption
End Property

Public Property Let BCaptionzwei(Schrift As String)
    m_Buttonzweicaption = Schrift
End Property

Public Property Get BCaptiondrei() As String
    BCaptionzwei = m_Buttondreicaption
End Property
Public Property Let BCaptiondrei(Schrift As String)
    m_Buttondreicaption = Schrift
End Property

Public Property Get BCaptionzwei() As String
    BCaptionzwei = m_Buttonzweicaption
End Property

Public Property Let Überschrift(Schrift As String)
    m_Überschrift = Schrift
End Property

Public Property Get Überschrift() As String
     Überschrift = m_Überschrift
End Property

Public Property Get Back() As Integer
    Back = m_Rückgabewert
End Property

Private Sub CancelButton_Click()
    m_Rückgabewert = 0
    Unload Me
    dlgAbfrage.Visible = False
End Sub

Private Sub cmdAusführen_Click()
    m_Rückgabewert = 1
    Unload Me
    dlgAbfrage.Visible = False

End Sub
Private Sub cmdWiederholen_Click()
    m_Rückgabewert = 2
    Unload Me
    dlgAbfrage.Visible = False

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    dlgAbfrage3.Icon = frmWKL00.Icon
    lblBeschriftung.Caption = m_Beschriftung
    CancelButton.Caption = m_Buttonzweicaption
    cmdAusführen.Caption = m_Buttoneinscaption
    cmdWiederholen.Caption = m_Buttondreicaption
    dlgAbfrage3.Caption = m_Überschrift
End Sub

