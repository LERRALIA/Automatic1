VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL28 
   BackColor       =   &H00FF0000&
   Caption         =   "Zahhlart"
   ClientHeight    =   7455
   ClientLeft      =   3645
   ClientTop       =   1815
   ClientWidth     =   10785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   7455
   ScaleWidth      =   10785
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command3 
      Height          =   615
      Index           =   6
      Left            =   5760
      TabIndex        =   10
      Top             =   6600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Abbrechen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   5055
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FF0000&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Euro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   492
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      Caption         =   "Zahlungsart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4815
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Bar"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Scheck"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
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
         Caption         =   "VISA"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
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
         Caption         =   "EuroCard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   4
         Left            =   240
         TabIndex        =   2
         Top             =   3240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
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
         Caption         =   "American Express"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   5
         Left            =   240
         TabIndex        =   1
         Top             =   3960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
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
         Caption         =   "EC-Karte"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin sevCommand3.Command Command3 
      Height          =   615
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Top             =   6600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Überweisung"
      Enabled         =   0   'False
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Wählen Sie eine Zahlart!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   5055
   End
End
Attribute VB_Name = "frmWKL28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sTraceNr    As String
    Dim sCent       As String

    Screen.MousePointer = 11
    Select Case Index
        Case Is = 0
            giZahlArt = 8       'Bar
            gcKreditKarte = "BA"
        Case Is = 1
            giZahlArt = 6       'Scheck
            gcKreditKarte = "SC"
        Case Is = 2
            giZahlArt = 17      'Karte
            gcKreditKarte = "VI"
        Case Is = 3
            giZahlArt = 17      'Karte
            gcKreditKarte = "EU"
        Case Is = 4
            giZahlArt = 17      'Karte
            gcKreditKarte = "AE"
        Case Is = 5
            giZahlArt = 17      'Karte
            gcKreditKarte = "EC"
            
        Case Is = 6 'abbruch
            giZahlArt = 0
            gcKreditKarte = ""
    End Select
    
    If giZahlArt = 17 Then
    
        If gbEcash Then
            Select Case gsEPartner
    
                Case Is = "ELP"
                
                    setzedrucker gcBonDrucker
        
                    lese_ELPAY_opt
                    
                    Label1(2).Caption = "Bedienen Sie jetzt das Kartenterminal!"
                    Label1(2).Refresh
                    
                    If CDbl(Label1(1).Caption) < 0 Then
                                                
                        'Storno
                        sCent = Label1(1).Caption
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_elPAY sTraceNr, sCent
                        
                        If giELPAY_Fehler > 0 Then
                        
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                 
                            Label1(2).Caption = "Fehler am Kartenterminal!"
                            Label1(2).Refresh
                            
                            setzedrucker gcListenDrucker
                            Exit Sub
                            'Abbruch
                        End If
                        
                    Else
                
                        'Zahlung
                        sCent = Label1(1).Caption
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_elPAY sCent
                        
                        If giELPAY_Fehler > 0 Then
                        
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                
                            Label1(2).Caption = "Fehler am Kartenterminal!"
                            Label1(2).Refresh
                            
                            setzedrucker gcListenDrucker
                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
                    setzedrucker gcListenDrucker
                    
                    
                    
                Case Is = "ZVT"
                
                    setzedrucker gcBonDrucker
        
                    lese_ZVT_opt
                    
                    Label1(2).Caption = "Bedienen Sie jetzt das Kartenterminal!"
                    Label1(2).Refresh
                    
                    If CDbl(Label1(1).Caption) < 0 Then
                                                
                        'Storno
                        sCent = Label1(1).Caption
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        dlgTaNr.Show 1
                    
                        sTraceNr = dlgTaNr.Back
                        
'                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "BNr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_ZVT sTraceNr
                        
                        If giZVT_Fehler > 0 Then
                        
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                 
                            Label1(2).Caption = "Fehler am Kartenterminal!"
                            Label1(2).Refresh
                            
                            setzedrucker gcListenDrucker
                            Exit Sub
                            'Abbruch
                        End If
                        
                    Else
                
                        'Zahlung
                        sCent = Label1(1).Caption
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_ZVT sCent
                        
                        If giZVT_Fehler > 0 Then
                        
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                
                            Label1(2).Caption = "Fehler am Kartenterminal!"
                            Label1(2).Refresh
                            
                            setzedrucker gcListenDrucker
                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
                    setzedrucker gcListenDrucker
                    
                Case Is = "ZV2"
                
'                    setzedrucker gcBonDrucker
        
                    lese_ZVT_opt2
                    
                    Label1(2).Caption = "Bedienen Sie jetzt das Kartenterminal!"
                    Label1(2).Refresh
                    
                    If CDbl(Label1(1).Caption) < 0 Then
                                                
                        'Storno
                        sCent = Label1(1).Caption
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        dlgTaNr.Show 1
                    
                        sTraceNr = dlgTaNr.Back
                        
'                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "BNr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_ZVT2 sTraceNr, sCent, False
                        
                        If giZVT2_Fehler > 0 Then
                        
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                 
                            Label1(2).Caption = "Fehler am Kartenterminal!"
                            Label1(2).Refresh
                            
'                            setzedrucker gcListenDrucker
                            Exit Sub
                            'Abbruch
                        End If
                        
                    Else
                
                        'Zahlung
                        sCent = Label1(1).Caption
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_ZVT2 sCent, False
                        
                        If giZVT2_Fehler > 0 Then
                        
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                
                            Label1(2).Caption = "Fehler am Kartenterminal!"
                            Label1(2).Refresh
                            
'                            setzedrucker gcListenDrucker
                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
                    
            End Select
        End If
    End If
    
    Unload frmWKL28
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Zahlungsart ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cFeld As String
    Dim dWert As Double
    Dim dSumme As Double
    
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    
    
'    frmWKL28.Top = Screen.Height / 2 - frmWKL28.Height / 2
'    frmWKL28.Left = Screen.Width / 2 - frmWKL28.Width / 2
    
    dSumme = 0
    dWert = 0
    
    For lcount = 1 To frmWKL24!MSFlexGrid2.Rows - 1
        frmWKL24!MSFlexGrid2.Row = lcount
        frmWKL24!MSFlexGrid2.Col = 0
        cFeld = frmWKL24!MSFlexGrid2.Text
        cFeld = Trim$(cFeld)
        If cFeld = "ausbuchen" Then
            frmWKL24!MSFlexGrid2.Col = 6
            cFeld = frmWKL24!MSFlexGrid2.Text
            cFeld = Trim$(cFeld)
            cFeld = fnMoveComma2Point$(cFeld)
            dWert = Val(cFeld)
            dSumme = dSumme + dWert
        End If
    Next lcount
    
    Label1(1).Caption = Format$(dSumme, "###,##0.00")
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Zahlungsart ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
