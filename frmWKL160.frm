VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWKL160 
   Caption         =   "Mailbox"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL160.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   11895
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   7
         Text            =   "frmWKL160.frx":0442
         Top             =   3960
         Width           =   9375
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   4
         Left            =   9480
         TabIndex        =   3
         Top             =   6840
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5741
         _Version        =   393216
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label1 
         Caption         =   "Nachricht"
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
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   7215
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   4
         Top             =   6840
         Width           =   7215
      End
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   1
      ToolTipText     =   "Hilfe"
      Top             =   360
      Width           =   375
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
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Mailbox"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmWKL160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
       
            
    
        Case 4
            Unload frmWKL160
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 11
            gsHelpstring = "Mailbox"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    WKL160Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Text3.Text = "Klicken Sie bitte in die Tabelle"
    
    If gbKL_LIVENACHRICHTEN = True Then
        If NewTableSuchenDBKombi("NACHRICHTEN", gdBase) = False Then
            CreateTableT3 "NACHRICHTEN", gdBase
        End If
        
        ImportiereNeueNachrichten
        Nachrichten_anzeigen
        
    Else
        anzeige "rot", "Möchten Sie die Nachrichten aus ihrer Zentrale empfangen, so nehmen Sie erst die erforderlichen Einstellungen unter: SERVICE/PROGRAMMEINSTELLUNGEN/Register Kisslive vor!", Label1(3)
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Nachrichten_anzeigen()
On Error GoTo LOKAL_ERROR

    FormatiereGridWKL160
    MoveNachrichten2GridWKL160

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Nachrichten_anzeigen"
    Fehler.gsFehlertext = "Im Programmteil Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FormatiereGridWKL160()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    With MSFlexGrid1
        .Rows = 10
        .Cols = 5
'        .FixedRows = 1
        .FixedCols = 0
        
        .Row = 0
        
        .Col = 0
        .ColWidth(0) = 1500
        .Text = "Datum"
        
        
        .Col = 1
        .ColWidth(1) = 1500
        .Text = "Uhrzeit"
        
        .Col = 2
        .ColWidth(2) = 5500
        .Text = "Betreff"
        
        .Col = 3
        .ColWidth(3) = 3000
        .Text = "Absender"
        
        .Col = 4
        .ColWidth(4) = 0
        .Text = "Nr"

    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereGridWKL160"
    Fehler.gsFehlertext = "Im Programmteil Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MoveNachrichten2GridWKL160()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow                As Long
    Dim lAnzRecords         As Long
    Dim cSQL                As String
    Dim rsrs                As DAO.Recordset
    Dim i                   As Integer
    Dim bGelesen            As Boolean
    
    lrow = 0
    
    cSQL = "Select * from Nachrichten Order by gelesen desc, Adate desc, Azeit desc"
    
    MSFlexGrid1.Redraw = False

    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
        
            bGelesen = False
        
            lAnzRecords = rsrs.RecordCount
            MSFlexGrid1.Rows = lAnzRecords + 1
            lrow = lrow + 1
           
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = 0
            If Not IsNull(rsrs!ADATE) Then
                MSFlexGrid1.Text = rsrs!ADATE
            Else
                MSFlexGrid1.Text = ""
            End If
           
            MSFlexGrid1.Col = 1
            
            If Not IsNull(rsrs!AZEIT) Then
                MSFlexGrid1.Text = rsrs!AZEIT
            Else
                MSFlexGrid1.Text = ""
            End If
            
            MSFlexGrid1.Col = 2
            
            If Not IsNull(rsrs!BETREFF) Then
                MSFlexGrid1.Text = rsrs!BETREFF
            Else
                MSFlexGrid1.Text = ""
            End If
            
            MSFlexGrid1.Col = 3
            
            If Not IsNull(rsrs!ABSENDER) Then
                MSFlexGrid1.Text = rsrs!ABSENDER
            Else
                MSFlexGrid1.Text = ""
            End If
            
            MSFlexGrid1.Col = 4
            
            If Not IsNull(rsrs!lfnr) Then
                MSFlexGrid1.Text = rsrs!lfnr
            Else
                MSFlexGrid1.Text = ""
            End If
            
            If Not IsNull(rsrs!gelesen) Then
                bGelesen = rsrs!gelesen
            Else
                bGelesen = False
            End If
            
            If bGelesen = False Then
                For i = 0 To MSFlexGrid1.Cols - 1
                
                    MSFlexGrid1.Col = i
                    MSFlexGrid1.CellFontBold = True
                   
                Next i
            Else
                For i = 0 To MSFlexGrid1.Cols - 1
                
                    MSFlexGrid1.Col = i
                    MSFlexGrid1.CellFontBold = False
                   
                Next i
            End If
            
            
            rsrs.MoveNext
        Loop
        
        MSFlexGrid1.Redraw = True
        

        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveNachrichten2GridWKL160"
    Fehler.gsFehlertext = "Im Programmteil Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKL160Positionieren()
On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 960
    Frame1.Left = 120
    Frame1.Width = 11775
    Frame1.Height = 7455
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL160Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo LOKAL_ERROR

    Dim lLFNR As Long

    MSFlexGrid1.Col = 4
    lLFNR = Val(MSFlexGrid1.Text)
    
    Text3.Text = ermNachrichtenMessage(lLFNR)
    
    'update Nachrichten auf gelesen
    Nachrichten_aufGelesen lLFNR
    
    Nachrichten_anzeigen
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
