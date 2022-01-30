VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKL56 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Mehrwertsteuersätze bestimmen"
   ClientHeight    =   5670
   ClientLeft      =   3225
   ClientTop       =   1560
   ClientWidth     =   11910
   Icon            =   "frmWKL56.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   5670
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   45
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8040
      MaxLength       =   2
      TabIndex        =   46
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8040
      MaxLength       =   2
      TabIndex        =   48
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   47
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton BTN_Etiketten 
      Caption         =   "Etiketten"
      Height          =   255
      Left            =   5640
      TabIndex        =   44
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton BTN_Runden 
      Caption         =   "Runden"
      Height          =   255
      Left            =   5640
      TabIndex        =   43
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BTN_KalkErm 
      Caption         =   "Kalkulieren"
      Height          =   255
      Left            =   9240
      TabIndex        =   35
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton BTN_KalkVoll 
      Caption         =   "Kalkulieren"
      Height          =   255
      Left            =   9240
      TabIndex        =   34
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton BTN_BackErm 
      Caption         =   "Rückgängig"
      Height          =   255
      Left            =   10560
      TabIndex        =   31
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton BTN_BackVoll 
      Caption         =   "Rückgängig"
      Height          =   255
      Left            =   10560
      TabIndex        =   30
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   2535
      Left            =   840
      TabIndex        =   14
      Top             =   3000
      Width           =   3855
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   12
         Left            =   2280
         TabIndex        =   27
         Top             =   1680
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   11
         Left            =   840
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "CE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   ","
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   9
         Left            =   3000
         TabIndex        =   24
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   8
         Left            =   2280
         TabIndex        =   23
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   7
         Left            =   1560
         TabIndex        =   22
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   6
         Left            =   840
         TabIndex        =   21
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   4
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   3
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   2
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   1
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Schließen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   735
      Index           =   0
      Left            =   2640
      TabIndex        =   12
      Top             =   2040
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Speichern"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBL_AnzeigeEtiketten 
      Caption         =   "Anzeige Etiketten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   52
      Top             =   5160
      Width           =   4575
   End
   Begin VB.Label LBL_AnzeigeRundung 
      Caption         =   "Anzeige Rundung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   51
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label LBL_Rundungsvariante 
      Caption         =   "Rundungsregel Variante 2 anwenden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   50
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label LBL_AnzeigeKalk 
      Caption         =   "Anzeige Kalkulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   49
      Top             =   2280
      Width           =   6135
   End
   Begin VB.Label Label12 
      Caption         =   $"frmWKL56.frx":0442
      Height          =   495
      Left            =   5640
      TabIndex        =   42
      Top             =   4200
      Width           =   6135
   End
   Begin VB.Label Label11 
      Caption         =   "Einstellung der Rundungsregel unter: Service/Programmeinstellungen/Register Voreinstellungen"
      Height          =   495
      Left            =   5640
      TabIndex        =   41
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label10 
      Caption         =   $"frmWKL56.frx":04E3
      Height          =   495
      Left            =   5640
      TabIndex        =   40
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label Label9 
      Caption         =   "%"
      Height          =   255
      Left            =   8640
      TabIndex        =   39
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "%"
      Height          =   255
      Left            =   8640
      TabIndex        =   38
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "auf"
      Height          =   255
      Left            =   8040
      TabIndex        =   37
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "von"
      Height          =   255
      Left            =   7560
      TabIndex        =   36
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "erm. MwSt - Artikel"
      Height          =   255
      Left            =   5640
      TabIndex        =   33
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "volle MwSt - Artikel"
      Height          =   255
      Left            =   5640
      TabIndex        =   32
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Kassenverkaufspreise anpassen"
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   120
      Width           =   3855
   End
   Begin VB.Line Line1 
      X1              =   5520
      X2              =   5520
      Y1              =   120
      Y2              =   5520
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "O"
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
      Index           =   8
      Left            =   2160
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ohne MWSt"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "E"
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
      Index           =   6
      Left            =   2160
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ermäßigte MWSt"
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
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "V"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "volle MWSt"
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MWSt-Satz in %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Abkürzung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mehrwertsteuer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmWKL56"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BTN_BackErm_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset

    Screen.MousePointer = 11

    sSQL = "Update ARTIKEL inner join KALKERMARTIKEL on Artikel.artnr = KALKERMARTIKEL.artnr SET ARTIKEL.KVKPR1 = KALKERMARTIKEL.KVKPR1 where Artikel.MWST = 'E'"
    gdBase.Execute sSQL, dbFailOnError

    Dim lCountKalkArtikel As Long
    lCountKalkArtikel = 0
    
    sSQL = "Select * from KALKERMARTIKEL"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lCountKalkArtikel = rsrs.RecordCount
    End If
    
    rsrs.Close
        
    anzeige "normal", "Kalkulation zurückgesetzt fertig!, " & lCountKalkArtikel & " Artikel(Mwst=E) wurden zurückgesetzt", LBL_AnzeigeKalk

    loeschNEW "KALKERMARTIKEL", gdBase
        
        
    'Buttonanzeige
    If NewTableSuchenDBKombi("KALKERMARTIKEL", gdBase) Then
        BTN_BackErm.Visible = True
        BTN_KalkErm.Visible = False
    Else
        BTN_BackErm.Visible = False
        BTN_KalkErm.Visible = True
    End If
    
    If NewTableSuchenDBKombi("KALKVOLLARTIKEL", gdBase) And NewTableSuchenDBKombi("KALKERMARTIKEL", gdBase) Then
        BTN_Etiketten.Enabled = True
        BTN_Runden.Enabled = True
    Else
        BTN_Etiketten.Enabled = False
        BTN_Runden.Enabled = False
    End If
        
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BTN_BackErm_Click"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub BTN_BackVoll_Click()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As DAO.Recordset

    Screen.MousePointer = 11

    sSQL = "Update ARTIKEL inner join KALKVOLLARTIKEL on Artikel.artnr = KALKVOLLARTIKEL.artnr SET ARTIKEL.KVKPR1 = KALKVOLLARTIKEL.KVKPR1 where Artikel.MWST = 'V'"
    gdBase.Execute sSQL, dbFailOnError

    Dim lCountKalkArtikel As Long
    lCountKalkArtikel = 0
    
    sSQL = "Select * from KALKVOLLARTIKEL"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lCountKalkArtikel = rsrs.RecordCount
    End If
    
    rsrs.Close
        
    anzeige "normal", "Kalkulation zurückgesetzt fertig!, " & lCountKalkArtikel & " Artikel(Mwst=V) wurden zurückgesetzt", LBL_AnzeigeKalk

    loeschNEW "KALKVOLLARTIKEL", gdBase
        
        
    'Buttonanzeige
    If NewTableSuchenDBKombi("KALKVOLLARTIKEL", gdBase) Then
        BTN_BackVoll.Visible = True
        BTN_KalkVoll.Visible = False
    Else
        BTN_BackVoll.Visible = False
        BTN_KalkVoll.Visible = True
    End If
    
    If NewTableSuchenDBKombi("KALKVOLLARTIKEL", gdBase) And NewTableSuchenDBKombi("KALKERMARTIKEL", gdBase) Then
        BTN_Etiketten.Enabled = True
        BTN_Runden.Enabled = True
    Else
        BTN_Etiketten.Enabled = False
        BTN_Runden.Enabled = False
    End If
        
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BTN_KalkVoll_Click"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub BTN_Etiketten_Click()
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset

    sSQL = "Delete from ETIDRU "
    gdBase.Execute sSQL, dbFailOnError

    'Etiketten abstellen
    
    sSQL = "Insert into etidru select artnr, bezeich, KVKPR1 as vkpr "
    sSQL = sSQL & ",bestand,bestand as anzahl, libesnr, lpz,ean,linr,1 as filnr, '' as pcname  from Artikel "
    sSQL = sSQL & " where Bestand > 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    Dim lCountEtiArtikel As Long
    lCountEtiArtikel = 0
    
    sSQL = "Select * from artikel where bestand > 0"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lCountEtiArtikel = rsrs.RecordCount
    End If
    rsrs.Close
    

    Dim lSumBestand As Long
    lSumBestand = 0
    
    sSQL = "Select sum(Bestand) as maxi from artikel where bestand > 0"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lSumBestand = rsrs!maxi
        End If
    End If
    
    rsrs.Close
    
    anzeige "normal", "Fertig!, " & lCountEtiArtikel & " verschiedene Strichcodeetiketten (Gesamtbestand von " & lSumBestand & ") in den Etikettenpool geschrieben", LBL_AnzeigeEtiketten
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BTN_Etiketten_Click"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub BTN_KalkErm_Click()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    If IsNumeric(Text3.Text) = False Then
        anzeige "rot", "nur Ziffern eingeben", LBL_AnzeigeKalk
        Exit Sub
    End If

    If IsNumeric(Text4.Text) = False Then
        anzeige "rot", "nur Ziffern eingeben", LBL_AnzeigeKalk
        Exit Sub
    End If
    Screen.MousePointer = 11
        
    MWST_SET

    loeschNEW "KALKERMARTIKEL", gdBase
    
    Dim dVonMwst As Double
    Dim dAufMwst As Double
    
    dVonMwst = 100 + CDbl(Text3.Text)
    dAufMwst = 100 + CDbl(Text4.Text)

    sSQL = "Select ARTNR,KVKPR1 into KALKERMARTIKEL from Artikel where MWST = 'E'"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update ARTIKEL SET KVKPR1 = KVKPR1 * " & dAufMwst & " /" & dVonMwst & "  where MWST = 'E'"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update ARTIKEL SET KVKPR1 = Round(KVKPR1,2) where MWST = 'E'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index ARTNR on KALKERMARTIKEL (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError
    
    Dim lCountKalkArtikel As Long
    lCountKalkArtikel = 0
    
    sSQL = "Select * from KALKERMARTIKEL"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lCountKalkArtikel = rsrs.RecordCount
    End If
    rsrs.Close
    
    anzeige "normal", "Kalkulation fertig!, " & lCountKalkArtikel & " Artikel(Mwst=E) wurden kalkuliert", LBL_AnzeigeKalk
    
    'Buttonanzeige
    If NewTableSuchenDBKombi("KALKERMARTIKEL", gdBase) Then
        BTN_BackErm.Visible = True
        BTN_KalkErm.Visible = False
    Else
        BTN_BackErm.Visible = False
        BTN_KalkErm.Visible = True
    End If
    
    If NewTableSuchenDBKombi("KALKVOLLARTIKEL", gdBase) And NewTableSuchenDBKombi("KALKERMARTIKEL", gdBase) Then
        BTN_Etiketten.Enabled = True
        BTN_Runden.Enabled = True
    Else
        BTN_Etiketten.Enabled = False
        BTN_Runden.Enabled = False
    End If
    

    Screen.MousePointer = 0


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BTN_KalkErm_Click"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub BTN_KalkVoll_Click()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset

    If IsNumeric(Text2.Text) = False Then
        anzeige "rot", "nur Ziffern eingeben", LBL_AnzeigeKalk
        Exit Sub
    End If

    If IsNumeric(Text5.Text) = False Then
        anzeige "rot", "nur Ziffern eingeben", LBL_AnzeigeKalk
        Exit Sub
    End If
    Screen.MousePointer = 11
        
    MWST_SET

    loeschNEW "KALKVOLLARTIKEL", gdBase
    
    Dim dVonMwst As Double
    Dim dAufMwst As Double

    dVonMwst = 100 + CDbl(Text2.Text)
    dAufMwst = 100 + CDbl(Text5.Text)

    sSQL = "Select ARTNR,KVKPR1 into KALKVOLLARTIKEL from Artikel where MWST = 'V'"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update ARTIKEL SET KVKPR1 = KVKPR1 * " & dAufMwst & " /" & dVonMwst & " where MWST = 'V'"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update ARTIKEL SET KVKPR1 = Round(KVKPR1,2) where MWST = 'V'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index ARTNR on KALKVOLLARTIKEL (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError
        
    Dim lCountKalkArtikel As Long
    
    lCountKalkArtikel = 0
    
    sSQL = "Select * from KALKVOLLARTIKEL"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lCountKalkArtikel = rsrs.RecordCount
    End If
    
    rsrs.Close
    
    anzeige "normal", "Kalkulation fertig!, " & lCountKalkArtikel & " Artikel(Mwst=V) wurden kalkuliert", LBL_AnzeigeKalk


    'Buttonanzeige
    If NewTableSuchenDBKombi("KALKVOLLARTIKEL", gdBase) Then
        BTN_BackVoll.Visible = True
        BTN_KalkVoll.Visible = False
    Else
        BTN_BackVoll.Visible = False
        BTN_KalkVoll.Visible = True
    End If
    
    If NewTableSuchenDBKombi("KALKVOLLARTIKEL", gdBase) And NewTableSuchenDBKombi("KALKERMARTIKEL", gdBase) Then
        BTN_Etiketten.Enabled = True
        BTN_Runden.Enabled = True
    Else
        BTN_Etiketten.Enabled = False
        BTN_Runden.Enabled = False
    End If
        
        
    Screen.MousePointer = 0


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BTN_KalkVoll_Click"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub BTN_Runden_Click()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    Dim lMaxDatensaetze As Long
    Dim lZaehlerDatensaetze As Long
    
    sSQL = "Select * from Artikel where MWST in ('V','E')"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMaxDatensaetze = rsrs.RecordCount
        lZaehlerDatensaetze = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            rsrs.Edit
            rsrs!KVKPR1 = Runden(CDbl(rsrs!KVKPR1))
            rsrs.Update
    
            rsrs.MoveNext
            
            anzeige "normal", "...noch " & lZaehlerDatensaetze & " Artikelpreise werden gerundet", LBL_AnzeigeRundung
            
            lZaehlerDatensaetze = lZaehlerDatensaetze - 1
            
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "Rundung fertig!, " & lMaxDatensaetze & " Artikelpreise wurden gerundet", LBL_AnzeigeRundung
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BTN_Runden_Click"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
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

Private Sub LeseMWStSaetzeWKL56()
    On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    cSQL = "Select * from MWSTSATZ WHERE FurJahr=" & Year(Date)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!VOLL) Then
            Text1(0).Text = rsrs!VOLL
        Else
            Text1(0).Text = ""
        End If
        If Not IsNull(rsrs!ERM) Then
            Text1(1).Text = rsrs!ERM
        Else
            Text1(1).Text = ""
        End If
        If Not IsNull(rsrs!OHNE) Then
            Text1(2).Text = rsrs!OHNE
        Else
            Text1(2).Text = 0
        End If
    Else
        Text1(0).Text = "0"
        Text1(1).Text = "0"
        Text1(2).Text = "0"
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseMWStSaetzeWKL56"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub BereiteMWSTKalkVor()
    On Error GoTo LOKAL_ERROR
    
    'Kalkulation
    If NewTableSuchenDBKombi("KALKVOLLARTIKEL", gdBase) Then
        BTN_BackVoll.Visible = True
        BTN_KalkVoll.Visible = False
    Else
        BTN_BackVoll.Visible = False
        BTN_KalkVoll.Visible = True
    End If
    
    If NewTableSuchenDBKombi("KALKERMARTIKEL", gdBase) Then
        BTN_BackErm.Visible = True
        BTN_KalkErm.Visible = False
    Else
        BTN_BackErm.Visible = False
        BTN_KalkErm.Visible = True
    End If

    If NewTableSuchenDBKombi("KALKVOLLARTIKEL", gdBase) And NewTableSuchenDBKombi("KALKERMARTIKEL", gdBase) Then
        BTN_Etiketten.Enabled = True
        BTN_Runden.Enabled = True
    Else
        BTN_Etiketten.Enabled = False
        BTN_Runden.Enabled = False
    End If

    'Rundungsregel

    If gbSPEZRU Then
    
        If gbSPEZVAR = 1 Then
            anzeige "normal", "Rundungsregel Variante 1 anwenden", LBL_Rundungsvariante
        ElseIf gbSPEZVAR = 2 Then
            anzeige "normal", "Rundungsregel Variante 2 anwenden", LBL_Rundungsvariante
        ElseIf gbSPEZVAR = 3 Then
            anzeige "normal", "Rundungsregel Variante 3 anwenden", LBL_Rundungsvariante
        ElseIf gbSPEZVAR = 4 Then
            anzeige "normal", "Rundungsregel Variante 4 anwenden", LBL_Rundungsvariante
        End If
        
    End If



Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BereiteMWSTKalkVor"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MWST_SET()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "MWSTMIST", gdBase
    
    sSQL = "Create table MWSTMIST"
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Artnr int "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into MWSTMIST"
    sSQL = sSQL & " Select Artnr from  Artikel "
    sSQL = sSQL & " where MWST not in ('V','E','O')"
    gdBase.Execute sSQL, dbFailOnError

    'Updaten Artikel
    sSQL = "Update Artikel set MWST = 'V' where artnr in(Select artnr from MWSTMIST ) "
    gdBase.Execute sSQL, dbFailOnError
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MWST_SET"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub SchreibeMWStSaetzeWKL56()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount  As Long
    Dim cSQL    As String
    Dim cVoll   As String
    Dim cErm    As String
    Dim cOhne   As String
    
    Dim dVoll   As Double
    Dim dErm    As Double
    Dim dOhne   As Double
    Dim rsrs    As Recordset
    
    cVoll = Text1(0).Text
    cVoll = fnMoveComma2Point$(cVoll)
    dVoll = Val(cVoll)
    
    cErm = Text1(1).Text
    cErm = fnMoveComma2Point$(cErm)
    dErm = Val(cErm)
    
    cOhne = Text1(2).Text
    cOhne = fnMoveComma2Point$(cOhne)
    dOhne = Val(cOhne)
    
    cSQL = "Select * from MWSTSATZ WHERE FurJahr=" & Year(Date)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
        rsrs!VOLL = dVoll
        rsrs!ERM = dErm
        rsrs!OHNE = dOhne
        rsrs.Update
    Else
        rsrs.AddNew
        rsrs!id = getNeuIdFurNeueMWST
        rsrs!FurJahr = Year(Date)
        rsrs!VOLL = dVoll
        rsrs!ERM = dErm
        rsrs!OHNE = dOhne
        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing
    
    RefreshGlobalVariableMWST
    MsgBox "MWSt-Sätze gespeichert!", vbInformation, "ERFOLG"
    
    Unload frmWKL56
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeMWStSaetzeWKL56"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub



Private Sub RefreshGlobalVariableMWST()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    cSQL = "Select * from MWSTSATZ WHERE FurJahr=" & Year(Date)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!VOLL) Then
            gdMWStV = rsrs!VOLL
        Else
            gdMWStV = 0
        End If
        If Not IsNull(rsrs!ERM) Then
            gdMWStE = rsrs!ERM
        Else
            gdMWStE = 0
        End If
        If Not IsNull(rsrs!OHNE) Then
            gdMWStO = rsrs!OHNE
        Else
            gdMWStO = 0
        End If
    Else
        gdMWStV = 0
        gdMWStE = 0
        gdMWStO = 0
        MsgBox "Es konnten keine MWST-Sätze gelesen werden!", vbCritical, "Winkiss Hinweis:"
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "frmWKL56"
    Fehler.gsFunktion = "RefreshGlobalVariableMWST"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub



Function getNeuIdFurNeueMWST() As Integer
On Error GoTo LOKAL_ERROR:
 
 Dim cmdSql As String
 Dim rsrsrs As Recordset
 
 cmdSql = "SELECT MAX(id)+1 as maxId FROM MWSTSATZ"
 
 Set rsrsrs = gdBase.OpenRecordset(cmdSql)
 If Not rsrsrs.EOF Then

      If Not IsNull(rsrsrs!maxId) Then
       getNeuIdFurNeueMWST = rsrsrs!maxId
      End If
      
 End If
 
Exit Function

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "getNeuIdFurNeueMWST"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function


Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
'    PositionierenWKL36
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    LeseMWStSaetzeWKL56
    
    BereiteMWSTKalkVor
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub SSCommand1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Select Case index
        Case Is = 0
            SchreibeMWStSaetzeWKL56
            
        Case Is = 1
            Unload frmWKL56
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub SSCommand2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iFeld As Integer
    
    iFeld = Val(Label2.Caption)
    
    Select Case index
        Case 0 To 9     'Ziffern
            Text1(iFeld).Text = Text1(iFeld).Text & SSCommand2(index).Caption
        Case Is = 10    'Komma
            If InStr(Text1(iFeld).Text, ",") = 0 Then
                Text1(iFeld).Text = Text1(iFeld).Text & SSCommand2(index).Caption
            End If
        Case Is = 11    'CE
            If Len(Text1(iFeld).Text) > 0 Then
                Text1(iFeld).Text = Left(Text1(iFeld).Text, Len(Text1(iFeld).Text) - 1)
            End If
        Case Is = 12    'C
            Text1(iFeld).Text = ""
    End Select
    
    Text1(iFeld).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Text1_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = glSelBack1
    Text1(index).SelStart = 0
    Text1(index).SelLength = Len(Text1(index).Text)
    Label2.Caption = Trim$(Str$(index))
    Label2.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Sub Text1_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil MWST ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


