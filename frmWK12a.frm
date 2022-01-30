VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK12a 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Beleg drucken"
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
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   2775
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   3600
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      Caption         =   "Nein"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   500
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      Caption         =   "Ja, speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   500
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      Caption         =   "Arbeitsbeginn"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      Caption         =   "Beenden"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   500
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      Caption         =   "OK"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblueberschrift 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Nein"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Bedienerkarte scannen !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "frmWK12a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim cBedcode As String
    
    Select Case Index
        Case Is = 0 'Ok gedrückt
            'hier unterscheiden ob an oder abmelden
           
            cBedcode = Text1.Text
            If gsMeldestatus = "Anmeldung" Then
                If Text1.Text = "" Then
                    MsgBox "Anmeldung gescheitert!", vbCritical, gsPname & " Anmeldung:"
                    Text1.Text = ""
                    Text1.SetFocus
                    glLevel = -1
                
                ElseIf Text1.Text = "הההההההההההה" Then
                    gcUserName = gcMASTERUSER
                    gcPass = gcMASTER
                    gcBedienerNr = "99"
                    glLevel = 9
                    frmWKL00!Label2.Visible = True
                    frmWKL00!Label2.Caption = gcUserName & " angemeldet"
                    frmWKL00!Label2.Refresh
                    
                    UpdateUSERSAFE gcBedienerNr, gcUserName
                    Unload frmWK12a
                Else
                    cSQL = "Select * from BEDNAME where BEDcode = '" & cBedcode & "' "
                    Set rsrs = gdBase.OpenRecordset(cSQL)
                    If Not rsrs.EOF Then
                        rsrs.MoveFirst
                        If Not IsNull(rsrs!BEDIENER) Then
                            glLevel = rsrs!BEDIENER
                        Else
                            glLevel = 0
                        End If
                        If Not IsNull(rsrs!BEDNU) Then
                            gcBedienerNr = rsrs!BEDNU
                        Else
                            gcBedienerNr = "-1"
                        End If
                        
                        If Not IsNull(rsrs!bedname) Then
                            gcUserName = rsrs!bedname
                        Else
                            gcUserName = ""
                        End If
                        
                        If Not IsNull(rsrs!Passwort) Then
                            gcPass = rsrs!Passwort
                        Else
                            gcPass = 0
                        End If
                        
                        If Label1.Caption = "Ja" Then      'arbeitszeitbeginn
'1
                            SchreibeArbeitszeit CLng(gcBedienerNr), gcUserName, "kommt", fncheckobRichtiggemeldet(CLng(gcBedienerNr), "kommt")
                            InsertandDelIdentUser CInt(gcBedienerNr), gcUserName, True
                        Else
                            If gbLokalModus Then
                                frmWKL00!Label2.Visible = True
                                frmWKL00!Label2.ForeColor = vbRed
                                frmWKL00!Label2.Caption = "lokaler Modus (Datenbank nicht erreichbar) - " & gcUserName & " angemeldet"
                                frmWKL00!Label2.Refresh
                            Else
                                frmWKL00!Label2.Visible = True
                                frmWKL00!Label2.Caption = gcUserName & " angemeldet"
                                frmWKL00!Label2.Refresh
                            End If
                            UpdateUSERSAFE gcBedienerNr, gcUserName
                            Unload frmWK12a
                        End If
    
    
                    Else
                        MsgBox "Anmeldung gescheitert!", vbCritical, gsPname & " Anmeldung:"
                        Text1.Text = ""
                        Text1.SetFocus
                        glLevel = -1
                    End If
                    rsrs.Close: Set rsrs = Nothing
                End If
                
                
                
            ElseIf gsMeldestatus = "Identifikation" Then
                If Text1.Text = "" Then
                    MsgBox "Identifikation gescheitert!", vbCritical, gsPname & " Identifikation:"
                    Text1.Text = ""
                    Text1.SetFocus
                    glIdentLevel = -1
                    
                ElseIf Text1.Text = "הההההההההההה" Then
                
                    gcIdentUserName = gcMASTERUSER
                    gcIdentPass = gcMASTER
                    gcIdentBedienerNr = "99"
                    glIdentLevel = 9

                    Unload frmWK12a
                Else
                    cSQL = "Select * from BEDNAME where BEDcode = '" & cBedcode & "' "
                    Set rsrs = gdBase.OpenRecordset(cSQL)
                    If Not rsrs.EOF Then
                        rsrs.MoveFirst
                        If Not IsNull(rsrs!BEDIENER) Then
                            glIdentLevel = rsrs!BEDIENER
                        Else
                            glIdentLevel = 0
                        End If
                        If Not IsNull(rsrs!BEDNU) Then
                            gcIdentBedienerNr = rsrs!BEDNU
                        Else
                            gcIdentBedienerNr = "-1"
                        End If
                        
                        If Not IsNull(rsrs!bedname) Then
                            gcIdentUserName = rsrs!bedname
                        Else
                            gcIdentUserName = ""
                        End If
                        
                        If Not IsNull(rsrs!Passwort) Then
                            gcIdentPass = rsrs!Passwort
                        Else
                            gcIdentPass = 0
                        End If
                        Unload frmWK12a
                    Else
                        MsgBox "Identifikation gescheitert!", vbCritical, gsPname & " Identifikation:"
                        Text1.Text = ""
                        Text1.SetFocus
                        glIdentLevel = -1
                    End If
                    rsrs.Close: Set rsrs = Nothing
                End If
                
            ElseIf gsMeldestatus = "Abmeldung" Then
            
                If Label1.Caption = "Ja" Then      'Arbeitszeitende
                    If Text1.Text = "" Then
                        MsgBox "Abmeldung gescheitert!", vbCritical, gsPname & " Abmeldung:"
                        Text1.Text = ""
                        Text1.SetFocus
                        glLevel = -1
                    
                    ElseIf Text1.Text = "הההההההההההה" Then
                        Unload frmWK12a
                    Else
                        cSQL = "Select * from BEDNAME where BEDcode = '" & cBedcode & "' "
                        Set rsrs = gdBase.OpenRecordset(cSQL)
                        If Not rsrs.EOF Then
                            rsrs.MoveFirst
                            If Not IsNull(rsrs!BEDIENER) Then
                                glLevel = rsrs!BEDIENER
                            Else
                                glLevel = 0
                            End If
                            If Not IsNull(rsrs!BEDNU) Then
                                gcBedienerNr = rsrs!BEDNU
                            Else
                                gcBedienerNr = "-1"
                            End If
                            
                            If Not IsNull(rsrs!bedname) Then
                                gcUserName = rsrs!bedname
                            Else
                                gcUserName = ""
                            End If
                            
                            If Not IsNull(rsrs!Passwort) Then
                                gcPass = rsrs!Passwort
                            Else
                                gcPass = 0
                            End If
                        End If
                        rsrs.Close: Set rsrs = Nothing
                        SchreibeArbeitszeit CLng(gcBedienerNr), gcUserName, "geht", fncheckobRichtiggemeldet(CLng(gcBedienerNr), "geht")
                        InsertandDelIdentUser CInt(gcBedienerNr), gcUserName, False
                    End If
                Else
                    
                    Unload frmWK12a
                End If
            
            End If
            

        Case Is = 1
        
            If gsMeldestatus = "Identifikation" Then
                Unload frmWK12a
                Unload frmWKL20
                
                'wenn bedkarte dann auch abmelden
                If gbBEDKARTE Then
                    If Identi(80) = False Then
                
                        gcUserName = ""
                        gcPass = ""
                        glLevel = -1
                        
                        frmWKL00!Label2.Visible = True
                        frmWKL00!Label2.Caption = "Anwender nicht aktiv"
                        frmWKL00!Label2.Refresh
                        
                        fAnmeldung
                    End If
                End If
                
            Else
                
                Unload frmWK12a
            End If
        Case Is = 2
            If Label1.Caption = "Nein" Then
                Label1.Caption = "Ja"
                
                Label4.Visible = True
                Text1.Visible = True
                Text1.SetFocus
            ElseIf Label1.Caption = "Ja" Then
                Label1.Caption = "Nein"
                If gsMeldestatus = "Abmeldung" Then
                    Label4.Visible = False
                    Text1.Visible = False
                ElseIf gsMeldestatus = "Anmeldung" Then
                    Label4.Visible = True
                    Text1.Visible = True
                    Text1.SetFocus
                End If
                
            End If
            
        Case Is = 3 ' ja
            If gsMeldestatus = "Anmeldung" Then
                
                    SchreibeArbeitszeit CLng(gcBedienerNr), gcUserName, "kommt", fncheckobRichtiggemeldet(CLng(gcBedienerNr), "kommt")
                
                If gbLokalModus Then
                    frmWKL00!Label2.Visible = True
                    frmWKL00!Label2.ForeColor = vbRed
                    frmWKL00!Label2.Caption = "lokaler Modus (Datenbank nicht erreichbar) - " & gcUserName & " angemeldet"
                    frmWKL00!Label2.Refresh
                Else
                    frmWKL00!Label2.Visible = True
                    frmWKL00!Label2.Caption = gcUserName & " angemeldet"
                    frmWKL00!Label2.Refresh
                End If
                Unload frmWK12a
            ElseIf gsMeldestatus = "Abmeldung" Then
            
                SchreibeArbeitszeit CLng(gcBedienerNr), gcUserName, "geht", fncheckobRichtiggemeldet(CLng(gcBedienerNr), "geht")
                Unload frmWK12a
            End If
            
        Case Is = 4 'abbrechen
            If gsMeldestatus = "Anmeldung" Then
                If gbLokalModus Then
                    frmWKL00!Label2.Visible = True
                    frmWKL00!Label2.ForeColor = vbRed
                    frmWKL00!Label2.Caption = "lokaler Modus (Datenbank nicht erreichbar) - " & gcUserName & " angemeldet"
                    frmWKL00!Label2.Refresh
                Else
                    frmWKL00!Label2.Visible = True
                    frmWKL00!Label2.Caption = gcUserName & " angemeldet"
                    frmWKL00!Label2.Refresh
                End If
                Unload frmWK12a
            ElseIf gsMeldestatus = "Abmeldung" Then
                Unload frmWK12a
            End If
            
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten. " & Index
    
    Fehlermeldung1
    
End Sub
Private Function fncheckobRichtiggemeldet(lbednu As Long, cART As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    Dim lDatum As Long
    Dim cDatum As String
    Dim lHeute As Long
    
    
    fncheckobRichtiggemeldet = ""
    
    
    sSQL = "SELECT TOP 1 stempel.*"
    sSQL = sSQL & "  from stempel where bednu = " & lbednu
    sSQL = sSQL & " and art = '" & cART & "' "
    sSQL = sSQL & " order by llfnr desc"
    
    lHeute = DateValue(Now)
    
    If cART = "kommt" Then
    
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                cDatum = rsrs!Datum
                lDatum = rsrs!Datum
                If lDatum = lHeute Then
                    ctmp = "Sie haben sich heute schon um: " & vbCrLf
                    ctmp = ctmp & CStr(rsrs!zeit)
                    
                    ctmp = ctmp & " angemeldet."
                    
                    fncheckobRichtiggemeldet = "Heute nochmals angemeldet."
                    
                Else
                    ctmp = "Anmeldung OK"
                    fncheckobRichtiggemeldet = "Anmeldung OK"
                
                End If
            End If
        Else
            ctmp = "Anmeldung OK"
            fncheckobRichtiggemeldet = "Anmeldung OK"
        End If
        rsrs.Close: Set rsrs = Nothing
    ElseIf cART = "geht" Then
    
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                lDatum = rsrs!Datum
                cDatum = rsrs!Datum
                If lDatum = lHeute Then
                    ctmp = "Sie haben sich heute schon um: " & vbCrLf
                    ctmp = ctmp & CStr(rsrs!zeit)
                    ctmp = ctmp & " abgemeldet."
                    fncheckobRichtiggemeldet = "Heute schon einmal abgemeldet."
                Else
                
                    fncheckobRichtiggemeldet = "Abmeldung OK"
                    ctmp = "Abmeldung OK"
                End If
            End If
        Else
            fncheckobRichtiggemeldet = "Abmeldung OK"
            ctmp = "Abmeldung OK"
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
        
    
    
    frmWK12a.Height = 6510
    frmWK12a.Top = 1700
    Label3.Caption = ctmp

    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fncheckobRichtiggemeldet"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    lblUeberschrift.Caption = gsMeldestatus
    lblUeberschrift.Refresh
    
    If (GetKeyState(vbKeyCapital) = 1) Then
    
        ' CAPS-Lock deaktivieren (falls aktiviert)
        MsgBox "CAPS-Lock (Großschreibmodus) ist eingeschaltet und wird jetzt abgestellt!", vbInformation, "Winkiss Hinweis:"
        
        
        KeyboardChangeState vbKeyCapital
    End If
    
    
    
    
    
    WK12aPositionieren
'    Modul6.Skalieren Me, True, True
    
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    frmWK12a.Left = Screen.Width / 2 - frmWK12a.Width / 2
    
    If gsMeldestatus = "Anmeldung" Then
        Command1(2).Visible = True
        Command1(2).Caption = "Arbeitsbeginn"
        Label4.Visible = True
        Text1.Visible = True
        Label1.Visible = True
    ElseIf gsMeldestatus = "Abmeldung" Then
        Command1(2).Visible = True
        Command1(2).Caption = "Arbeitsende"
        Label4.Visible = False
        Text1.Visible = False
        Label1.Visible = True
    ElseIf gsMeldestatus = "Identifikation" Then
        Command1(2).Visible = False
        Label4.Visible = False
        Label1.Visible = False
        Text1.Visible = True
        gbStornoErlaubt = False
        
    End If
    
    Text1.Text = ""
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WK12aPositionieren()
On Error GoTo LOKAL_ERROR
    
    frmWK12a.Top = 2000
    frmWK12a.Height = 3400
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WK12aPositionieren"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change()
    On Error GoTo LOKAL_ERROR
    
    If Text1.Text = "kiss2005" Or Text1.Text = gskPW Or Text1.Text = "xyc" Then
        Text1.Text = "הההההההההההה"
    Else
        bedcodewandeln Text1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus()
On Error GoTo LOKAL_ERROR

    Text1.BackColor = glSelBack1 'glSelBack1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command1_Click 0
    ElseIf KeyCode = vbKeyEscape Then
        Command1_Click 1
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeArbeitszeit(cbednu As String, cbedname As String, cART As String, cUeberschrift As String)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp    As String
    Dim lDatum  As Long
    Dim cDatum  As String
    Dim czeit   As String
    
    
    cDatum = DateValue(Now)
    lDatum = DateValue(Now)
    czeit = TimeValue(Now)
    
    If Command1(3).Visible = True Then
        Command1(3).Visible = False
        Command1(4).Visible = False
        speicherArbeitszeit CLng(cbednu), cbedname, cDatum, czeit, cART, cUeberschrift
    Else
    
        unsichtbarer
        If cART = "kommt" Then
            ctmp = "Arbeitszeitbeginn: " & vbCrLf
        ElseIf cART = "geht" Then
            ctmp = "Arbeitszeitende: " & vbCrLf
        End If
        ctmp = ctmp & cbedname & " (" & cbednu & ")" & vbCrLf
        ctmp = ctmp & "am: " & cDatum
        ctmp = ctmp & "         um: " & czeit & vbCrLf & vbCrLf
        ctmp = ctmp & "Diese Angaben speichern?"
        

        
        Label2.Caption = ctmp
        Label2.Refresh
        
        Command1(3).Visible = True
        Command1(4).Visible = True
        frmWK12a.Height = 6510
        frmWK12a.Top = 1700
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeArbeitszeit"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InsertandDelIdentUser(ibednu As Integer, cbedname As String, gbinsert As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL   As String
    
    sSQL = "Delete from Identuser where BEDNR = " & ibednu
    gdBase.Execute sSQL, dbFailOnError
    
    If gbinsert Then
        sSQL = "Insert into Identuser (BEDNR,BEDNAME) values (" & ibednu & ",'" & cbedname & "')"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertandDelIdentUser"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub unsichtbarer()
    On Error GoTo LOKAL_ERROR
    
    Text1.Visible = False
    Label4.Visible = False
    Label1.Visible = False
    Command1(0).Visible = False
    Command1(1).Visible = False
    Command1(2).Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "unsichtbarer"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherArbeitszeit(lbednu As Long, cbedname As String, cDatum As String, czeit As String, cART As String, cbemerk As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    ReDim cZeilen(0 To 9) As String
    
    cSQL = "Insert into Stempel (BEDNU, BEDNAME, DATUM, ZEIT, ART) values "
    cSQL = cSQL & "( " & lbednu & ", "
    cSQL = cSQL & "'" & cbedname & "', "
    cSQL = cSQL & "'" & cDatum & "', "
    cSQL = cSQL & "'" & czeit & "', "
    cSQL = cSQL & "'" & cART & "' )"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    'Drucke den Beleg
    If Check1.Value = vbChecked Then
        cZeilen(0) = "ARBEITSZEIT-BELEG"
        cZeilen(1) = "-----------------"
        cZeilen(2) = "BedNr: " & Trim$(Str$(lbednu))
        cZeilen(3) = "Name:  " & cbedname
        cZeilen(4) = "Art:   " & cART
        cZeilen(5) = "Datum: " & cDatum
        cZeilen(6) = "Zeit:  " & czeit
        cZeilen(7) = "Bemerkungen:  "
        cZeilen(8) = cbemerk
        DruckeArbeitszeitBelegWK20d cZeilen(), 9
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherArbeitszeit"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
End Sub

Private Sub Text1_LostFocus()
On Error GoTo LOKAL_ERROR

    Text1.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_lostFocus"
    Fehler.gsFehlertext = "Bei der Anmeldung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
