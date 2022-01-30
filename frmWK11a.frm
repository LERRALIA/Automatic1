VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWK11a 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   3600
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4650
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWK11a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4650
   Begin sevCommand3.Command Command1 
      Height          =   255
      Left            =   80
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
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
      Caption         =   "Ansicht wechseln"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _Version        =   524288
      _ExtentX        =   8281
      _ExtentY        =   5741
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2004
      Month           =   5
      Day             =   3
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   255
      FirstDay        =   2
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.MonthView MV 
      Height          =   3510
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   6191
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   390266882
      CurrentDate     =   38427
   End
End
Attribute VB_Name = "frmWK11a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
    On Error GoTo LOKAL_ERROR
    
    gsDatum = Calendar1.Value
    Unload frmWK11a
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Calendar1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kalender auf. "
    Fehlermeldung1
End Sub

Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR

    If Calendar1.Visible = True Then
    
        Calendar1.Visible = False
        MV.Visible = True
        
        sSQL = "Update WKEINSTE Set mv = True"
        gdApp.Execute sSQL, dbFailOnError
        
        gbmv = True
    Else
    
        Calendar1.Visible = True
        MV.Visible = False
        
        sSQL = "Update WKEINSTE Set mv = False"
        gdApp.Execute sSQL, dbFailOnError
        
        gbmv = False
    End If
    
    WK11aPositionieren
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kalender auf. "
    Fehlermeldung1
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    WK11aPositionieren
'    Modul6.Schrift Me
    
'    Frm.Controls(i).DayFont.name = gsFont
'    Frm.Controls(i).GridFont.name = gsFont
    Calendar1.DayFont.Size = 12
    Calendar1.GridFont.Size = 12

'    Frm.Controls(i).TitleFont.name = gsFont
    Calendar1.TitleFont.Size = 12
    
    If gsDatum = "" Then
        Calendar1.Value = DateValue(Now)
        MV.Value = DateValue(Now)
    Else
        Calendar1.Value = gsDatum
        MV.Value = gsDatum
    End If
    
    gsDatum = ""
    
    If gbmv Then
        Calendar1.Visible = False
        MV.Visible = True
    Else
        Calendar1.Visible = True
        MV.Visible = False
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kalender auf. "
    Fehlermeldung1
End Sub
Private Sub WK11aPositionieren()
On Error GoTo LOKAL_ERROR

    If Val(glKalLeft) = 0 Then
        Me.Left = 0
    Else
        Me.Left = glKalLeft
    End If
    
    If Val(glKalTop) = 0 Then
        Me.Top = 0
    Else
        Me.Top = glKalTop
    End If
    
    If gbmv Then
        Me.Width = 3210
        Me.Height = 3950
        
        With Command1
            .Top = 3560
            .Left = 80
            .Width = 1815
            .Height = 255
        End With
    Else
        Me.Width = 4710
        Me.Height = 3600 '3225
        
        With Command1
            .Top = 3200
            .Left = 80
            .Width = 1815
            .Height = 255
        End With
    End If
    
    With Calendar1
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = 3225
    End With
    
    With MV
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = 3585
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WK11aPositionieren"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kalender auf. "
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    glKalTop = 0
    glKalLeft = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kalender auf. "
    
    Fehlermeldung1
End Sub
Private Sub MV_DateClick(ByVal DateClicked As Date)
On Error GoTo LOKAL_ERROR
    
    gsDatum = MV.Value
    Unload frmWK11a
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MV_DateClick"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kalender auf. "
    Fehlermeldung1
End Sub
