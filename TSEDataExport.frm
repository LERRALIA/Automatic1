VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TSEDataExport 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "TSE Data Export"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btnExport 
      Caption         =   "Export starten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Exportoptionen"
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7455
      Begin VB.TextBox txtTransNrBis 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox txtTransNrVon 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1500
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   270
         Left            =   3600
         TabIndex        =   12
         Top             =   2520
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   112001025
         CurrentDate     =   44314
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   270
         Left            =   1200
         TabIndex        =   9
         Top             =   2520
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   112001025
         CurrentDate     =   44314
      End
      Begin VB.OptionButton optZeit 
         BackColor       =   &H00C0C000&
         Caption         =   "in Zeitraum"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton optTransNr 
         BackColor       =   &H00C0C000&
         Caption         =   "nach Transactionsnummern"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin VB.OptionButton optGesamt 
         BackColor       =   &H00C0C000&
         Caption         =   "Gesamtes Archiv"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Ende :"
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
         Left            =   3000
         TabIndex        =   11
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Start :"
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
         Left            =   600
         TabIndex        =   10
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Ende :"
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
         Left            =   3000
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Start :"
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
         Left            =   600
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.CommandButton btnDurchsuchen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtExportPfad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label lblExportStatus 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   4440
      Width           =   5535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ausgabeverzeichnis :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "TSEDataExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDurchsuchen_Click()

 'Ausgabeverzeichnis auswählen
 ChooseFile.Left = (Me.Left + 1100)
 ChooseFile.Top = (Me.Top + 800)
 ChooseFile.Show 1
  
  
 If Trim(gbTSEExportPfad) = "" Then
   
     btnExport.Enabled = False
     Frame1.Visible = False
    Else
        
         
         If txtExportPfad.Text <> gbTSEExportPfad Then
            
            'der alte Pfad überschreiben
            Dim sSQL As String
            sSQL = "UPDATE TSESettings SET TSE_ExportPfad='" & gbTSEExportPfad & "'"
            gdApp.Execute sSQL, dbFailOnError
         
         End If
         
         txtExportPfad.Text = gbTSEExportPfad
         btnExport.Enabled = True
         Frame1.Visible = True
     
 End If
 
 
 
End Sub

Private Sub btnExport_Click()

 btnExport.Enabled = False

 lblExportStatus.Caption = ""
 lblExportStatus.Refresh
 
 txtTransNrVon.BackColor = vbWhite
 txtTransNrBis.BackColor = vbWhite
 
 Dim istWasSchonGewaehlt As Boolean
 
 '''''''''''''''''''''''''''''' Gesamtes Archiv ''''''''''''''''''''''''''
 If optGesamt.value Then                                                 '
    istWasSchonGewaehlt = True                                           '
    DatenExportieren                                                     '
    Exit Sub                                                             '
                                                                         '
 End If                                                                  '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 '''''''''''''''''''''''''''''' nach Transactions Nr '''''''''''''''''''''''''''''''''''
 If optTransNr.value Then                                                              '
                                                                                       '
    istWasSchonGewaehlt = True                                                         '
                                                                                       '
    If Trim(txtTransNrVon.Text) = "" Then                                              '
       txtTransNrVon.BackColor = vbRed                                                 '
                                                                                       '
    ElseIf Trim(txtTransNrBis.Text) = "" Then                                          '
        txtTransNrBis.BackColor = vbRed                                                '
                                                                                       '
    Else                                                                               '
        DatenExportierenNachTransNr CLng(txtTransNrVon.Text), CLng(txtTransNrBis.Text) '
        Exit Sub                                                                       '
    End If                                                                             '
                                                                                       '
 End If                                                                                '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 '''''''''''''''''''''''''''''' nach Zeitraum ''''''''''''''''''''''''''''
 If optZeit.value Then                                                   '
    istWasSchonGewaehlt = True                                           '
    DatenExportierenNachZeitraum DTPicker1.value, DTPicker2.value        '
    Exit Sub                                                             '
                                                                         '
 End If                                                                  '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 If Not istWasSchonGewaehlt Then
   MsgBox ("Bitte eine Exportoption auswählen !!!")
 End If
   
 btnExport.Enabled = True
 
End Sub
 

Private Sub Form_Load()
 
  If Trim(gbTSEExportPfad) = "" Then
  
     btnExport.Enabled = False
     Frame1.Visible = False
    Else
    
     txtExportPfad.Text = gbTSEExportPfad
     btnExport.Enabled = True
     Frame1.Visible = True
  End If
 
End Sub

Private Sub optGesamt_Click()
 DTPicker1.Enabled = False
 DTPicker2.Enabled = False
 txtTransNrVon.Enabled = False
 txtTransNrBis.Enabled = False
End Sub

Private Sub optTransNr_Click()
 DTPicker1.Enabled = False
 DTPicker2.Enabled = False
 
 txtTransNrVon.Enabled = True
 txtTransNrBis.Enabled = True
 
End Sub

Private Sub optZeit_Click()
 DTPicker1.Enabled = True
 DTPicker2.Enabled = True
 
 txtTransNrVon.Enabled = False
 txtTransNrBis.Enabled = False
 
End Sub

Private Sub txtTransNrBis_Change()

 txtTransNrBis.BackColor = vbWhite
  
 Dim textval As String
  
 textval = Trim(txtTransNrBis.Text)
 textval = Replace(textval, ".", "")
 textval = Replace(textval, ",", "")
 
  If IsNumeric(textval) Then
      txtTransNrBis.Text = CStr(textval)
    Else
      txtTransNrBis.Text = ""
    
  End If
  
End Sub

Private Sub txtTransNrVon_Change()

 txtTransNrVon.BackColor = vbWhite

 Dim textval As String
  
 textval = Trim(txtTransNrVon.Text)
 textval = Replace(textval, ".", "")
 textval = Replace(textval, ",", "")
 
  If IsNumeric(textval) Then
      txtTransNrVon.Text = CStr(textval)
    Else
      txtTransNrVon.Text = ""
    
  End If
  
End Sub
