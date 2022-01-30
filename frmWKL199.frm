VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL199 
   Caption         =   "Dewas Import"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL199.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtStartdatum 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Text            =   "10.01.2015"
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   9315
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   7320
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   6
      Left            =   9600
      TabIndex        =   3
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
      Caption         =   "importieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
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
   Begin MSComDlg.CommonDialog cdlopen 
      Left            =   6720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Startdatum(Tabelle Umsatz und Kassjour)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "- Dewas-Daten als Textdateien"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   10095
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Was benötigen Sie für den Datenimport?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11535
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   9255
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Dewas Import"
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
      TabIndex        =   1
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmWKL199"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad As String
    Dim sdbPfad As String

    Select Case Index
        Case 0
            Unload frmWKL199
        Case 6
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Wo sind die Dewas - Dateien?"

                .Filter = "DBUmAb*"
                .ShowSave

                sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
            End With
            
            DewasImport Label1(4), sPfad, txtStartdatum.Text
        
    End Select

err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Dewas Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub DewasImport(lblx As Label, cPfad As String, sStartdatum As String)
On Error GoTo LOKAL_ERROR

    lblx.Caption = TimeValue(Now) & ": Tabellen werden in den Speicher geladen...": lblx.Refresh
    
    Dim lPosEnde        As Long
    Dim cEinzelsatz     As String
    Dim lLenfil         As Long
    Dim lposSemi        As Long
    Dim lposSemiEnde    As Long
    Dim cWert           As String
    Dim lfnr1           As Long
    Dim lPos            As String
    Dim iFileNr         As Integer
    Dim cSatz1          As String
    Dim dWert           As Double
    Dim lcount          As Long
    Dim sSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim sArtnr          As String
    Dim sEAN            As String
    Dim j               As Integer
    Dim lLenVerbleib    As Long
    Dim slibesnr        As String
    Dim sVPE            As String
    Dim sLEKPR          As String
    
    '1. Tabelle Umsatz wird aufgefüllt
    
    If FileExists(cPfad & "\dbUmab") Then
    
        Dim sUmsKZ As String
        Dim sUmsDatum As String
        
        Dim sUmsG1 As String
        Dim sUmsV1 As String
        Dim sUmsE1 As String
        Dim sUmsO1 As String
        Dim sUmsKunz1 As String
        Dim sUmsEKPR1 As String
        Dim sUmsKRED1 As String
        
        Dim cFiliale As String
        
        lPos = 1
        lPosEnde = 1
        lposSemiEnde = 1
        
'         sSQL = "Delete from Umsatz "
'                        gdBase.Execute sSQL, dbFailOnError


       




    
        iFileNr = FreeFile
        Open cPfad & "\dbUmab" For Binary As #iFileNr
        If LOF(iFileNr) > 0 Then
        
            cSatz1 = Space$(LOF(iFileNr))
            Get #iFileNr, 1, cSatz1
        
            lLenfil = Len(cSatz1)
            lPosEnde = InStr(lPos, cSatz1, vbCrLf)
            lPos = lPos + lPosEnde - lPos + 2 'Kopfzeile überspringen
            
            lcount = 0
            
            Do
                lcount = lcount + 1
                
                lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
                lPos = lPos + lPosEnde - lPos + 2
                lposSemi = 1
                
                sUmsKZ = ""
                sUmsDatum = ""
                
                sUmsG1 = "0"
                sUmsV1 = "0"
                sUmsE1 = "0"
                sUmsO1 = "0"
                sUmsKunz1 = "0"
                sUmsEKPR1 = "0"
                sUmsKRED1 = "0"
    
    
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                '1. überspringen
                cFiliale = cWert
    
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                '2. überspringen
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                '3. Kennzeichen T = Tag
                sUmsKZ = cWert
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                '4. Datum
                sUmsDatum = cWert
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                
                cWert = SwapStr(cWert, ".", ",")
                sUmsEKPR1 = Format(cWert, "#####0.00")
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                cWert = SwapStr(cWert, ".", ",")
                sUmsG1 = Format(cWert, "#####0.00")
                
                For i = 0 To 6
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                Next i
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sUmsKunz1 = Val(cWert)
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
    
                If sUmsKZ = "T" Then
                    If DateValue(sUmsDatum) >= DateValue(sStartdatum) Then
                    
                    
                        sSQL = "Delete from Umsatz where Datum = " & CLng(DateValue(sUmsDatum))
                        gdBase.Execute sSQL, dbFailOnError

                        sSQL = "Insert into Umsatz "
                        sSQL = sSQL & " ( "
                        sSQL = sSQL & " Datum  "
                        sSQL = sSQL & ", UMSG1 "
                        sSQL = sSQL & ", UMSV1 "
                        sSQL = sSQL & ", UMSE1 "
                        sSQL = sSQL & ", UMSO1 "
                        sSQL = sSQL & ", Kunz1 "
                        sSQL = sSQL & ", EKPR1 "
                        sSQL = sSQL & ", Kred1 "
                        sSQL = sSQL & " ) "
                        sSQL = sSQL & " values "
                        sSQL = sSQL & " ( " & CLng(DateValue(sUmsDatum)) & ""
                        sSQL = sSQL & ", '" & sUmsG1 & "' "
                        sSQL = sSQL & ", '" & sUmsV1 & "' "
                        sSQL = sSQL & ", '" & sUmsE1 & "' "
                        sSQL = sSQL & ", '" & sUmsO1 & "' "
                        sSQL = sSQL & ", '" & sUmsKunz1 & "' "
                        sSQL = sSQL & ", '" & sUmsEKPR1 & "' "
                        sSQL = sSQL & ", '" & sUmsKRED1 & "' "
                        sSQL = sSQL & " ) "
                        gdBase.Execute sSQL, dbFailOnError
                        
                        
                        
                        
                                  
                    
                    End If
                End If
            Loop While lLenfil >= lPos
            
        End If
    End If
    
    'Ende
    '1. Tabelle Umsatz wird aufgefüllt _ Ende
    
    lblx.Caption = TimeValue(Now) & ": Lieferanten werden übernommen...": lblx.Refresh
    
    
    '2. Tabelle Lieferanten wird erstellt
    
    sSQL = "Delete from LISRT "
    gdBase.Execute sSQL, dbFailOnError
    
    If FileExists(cPfad & "\dblfts") Then
    
        Dim sLinr As String
        Dim sLinrBEZ As String
        Dim sLinrKUERZEL As String
        Dim sLinrORT As String
        Dim sLinrPLZ As String
        Dim sLinrStrasse As String
        Dim sLinrTEL As String
        Dim sLinrFAX As String
        Dim sLinrKUNDNR As String
        Dim sLinrMail As String
        Dim sLinrBestellMail As String
        Dim sLinrBestellGLN As String
        
        lPos = 1
        lPosEnde = 1
        lposSemiEnde = 1
        
        iFileNr = FreeFile
        Open cPfad & "\dblfts" For Binary As #iFileNr
        If LOF(iFileNr) > 0 Then
        
            cSatz1 = Space$(LOF(iFileNr))
            Get #iFileNr, 1, cSatz1
        
            lLenfil = Len(cSatz1)
            lPosEnde = InStr(lPos, cSatz1, vbCrLf)
            lPos = lPos + lPosEnde - lPos + 2 'Kopfzeile überspringen
            
            lcount = 0
            
            Do
                lcount = lcount + 1
                
                lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
                lPos = lPos + lPosEnde - lPos + 2
                lposSemi = 1
                
                sLinr = ""
                sLinrBEZ = ""
                sLinrKUERZEL = ""
                sLinrORT = ""
                sLinrPLZ = ""
                sLinrStrasse = ""
                sLinrTEL = ""
                sLinrFAX = ""
                sLinrKUNDNR = ""
                sLinrMail = ""
                sLinrBestellMail = ""
                sLinrBestellGLN = ""
    
    
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                '1. Linr
                sLinr = cWert
    
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                
                sLinrKUERZEL = cWert
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                
                sLinrBEZ = cWert
                sLinrBEZ = SwapStr(sLinrBEZ, "'", "")
                
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                
                sLinrStrasse = cWert
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sLinrPLZ = cWert
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sLinrORT = cWert
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sLinrTEL = cWert
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sLinrFAX = cWert
                '40
                
                For i = 0 To 39
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                Next i
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sLinrKUNDNR = cWert
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sLinrGLN = cWert
                
                For i = 0 To 28
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                Next i
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sLinrBestellMail = cWert
                
                For i = 0 To 26
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                Next i
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sLinrMail = cWert
                
                lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
    
                sSQL = "Insert into LISRT "
                sSQL = sSQL & " ( "
                sSQL = sSQL & " LINR  "
                sSQL = sSQL & ", LIEFBEZ "
                sSQL = sSQL & ", KUERZEL "
                sSQL = sSQL & ", STRASSE "
                sSQL = sSQL & ", STADT "
                sSQL = sSQL & ", PLZ "
                sSQL = sSQL & ", TEL "
                sSQL = sSQL & ", FAX "
                sSQL = sSQL & ", KUNDNR "
                sSQL = sSQL & ", EMAIL "
                sSQL = sSQL & ", adress "
                sSQL = sSQL & ", GLN "
                sSQL = sSQL & " ) "
                sSQL = sSQL & " values "
                sSQL = sSQL & " ( " & sLinr & ""
                sSQL = sSQL & ", '" & sLinrBEZ & "' "
                sSQL = sSQL & ", '" & sLinrKUERZEL & "' "
                sSQL = sSQL & ", '" & sLinrStrasse & "' "
                sSQL = sSQL & ", '" & sLinrORT & "' "
                sSQL = sSQL & ", '" & sLinrPLZ & "' "
                sSQL = sSQL & ", '" & sLinrTEL & "' "
                sSQL = sSQL & ", '" & sLinrFAX & "' "
                sSQL = sSQL & ", '" & sLinrKUNDNR & "' "
                sSQL = sSQL & ", '" & sLinrMail & "' "
                sSQL = sSQL & ", '" & sLinrBestellMail & "' "
                sSQL = sSQL & ", '" & sLinrGLN & "' "
                sSQL = sSQL & " ) "
                sSQL = SwapStr(sSQL, Chr(34), "")
                
                gdBase.Execute sSQL, dbFailOnError

            Loop While lLenfil >= lPos
            
        End If
    
    End If
    
    'Ende
    '2. Tabelle Lieferanten wird erstellt _ Ende
    
    
    
    lblx.Caption = TimeValue(Now) & ": EANs werden übernommen...": lblx.Refresh
    
    'Artean erstellen
    picprogress.Visible = True
    If FileExists(cPfad & "\dbReFe") Then

        loeschNEW "ARTEAN_DEWAS", gdBase

        sSQL = "Create Table ARTEAN_DEWAS"
        sSQL = sSQL & " ( "
        sSQL = sSQL & " ARTNR double "
        sSQL = sSQL & ", EAN Text(13) "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError

        sSQL = "select * from ARTEAN_DEWAS "
        Set rsrs = gdBase.OpenRecordset(sSQL)

        lPos = 1
        lPosEnde = 1
        lposSemiEnde = 1

        iFileNr = FreeFile
        Open cPfad & "\dbReFe" For Binary As #iFileNr
        If LOF(iFileNr) > 0 Then

            cSatz1 = Space$(LOF(iFileNr))
            Get #iFileNr, 1, cSatz1

            lLenfil = Len(cSatz1)
            lPosEnde = InStr(lPos, cSatz1, vbCrLf)
            lPos = lPos + lPosEnde - lPos + 2 'Kopfzeile überspringen

            lcount = 0

            Do
                lcount = lcount + 1

                lLenVerbleib = lLenfil - lPos

                txtStatus.Text = (lLenVerbleib * 100) / lLenfil

                j = lcount Mod 100
                If j = 0 Then
                    lblx.Caption = TimeValue(Now) & " " & lcount & "..."
                    lblx.Refresh
                End If

                lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
                lPos = lPos + lPosEnde - lPos + 2
                lposSemi = 1

                sArtnr = ""
                sEAN = ""

                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sEAN = Val(Left(cWert, 12))

                lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                sArtnr = cWert

                lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

                rsrs.AddNew
                rsrs!artnr = Val(sArtnr)
                rsrs!EAN = Trim(sEAN) & fn_errechne_Prüfziffer(Trim(sEAN))
                rsrs.Update

            Loop While lLenfil >= lPos

        End If

        rsrs.Close
    End If
    picprogress.Visible = False
    'Ende Artean


    lblx.Caption = TimeValue(Now) & ": Tabelle Artlief wird erstellt...": lblx.Refresh

    'Artlief erstellen etwas schneller
    
'    Dim slokalPfad As String
'     slokalPfad = gcDBPfad
'    If Right(cPfad, 1) <> "\" Then
'        cPfad = cPfad & "\"
'    End If
    
    loeschNEW "ARTLIEF_DEWAS", gdBase
        
    sSQL = "Create Table ARTLIEF_DEWAS"
    sSQL = sSQL & " ( "
    sSQL = sSQL & " ARTNR double "
    sSQL = sSQL & ", Linr long "
    sSQL = sSQL & ", libesnr Text(13) "
    sSQL = sSQL & ", VPE long"
    sSQL = sSQL & ", LEKPR double"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ARTLIEF_DEWAS Select "
    sSQL = sSQL & " dArtnr as ARTNR  "
    sSQL = sSQL & ", lLiefnr as Linr  "
    sSQL = sSQL & ", left(trim(cartnrl),13) as libesnr  "
    sSQL = sSQL & ", sibeh as VPE "
    sSQL = sSQL & ", dekl as LEKPR "
    sSQL = sSQL & " from [;DATABASE=" & cPfad & "\safe.MDB].dblftb"
    gdBase.Execute sSQL, dbFailOnError
    
    'Ende Artlief erstellen etwas schneller
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour_Dewas wird erstellt...": lblx.Refresh
    
    'Kassjour_Dewas erstellen
    
    loeschNEW "Kassjour_Dewas", gdBase
    
    sSQL = "Create Table Kassjour_Dewas"
    sSQL = sSQL & "("
    sSQL = sSQL & " ARTNR LONG "
    sSQL = sSQL & ", BEZEICH TEXT(35) "
    sSQL = sSQL & ", MENGE INTEGER "
    sSQL = sSQL & ", PREIS SINGLE "
    sSQL = sSQL & ", ADATE DATETIME "
    sSQL = sSQL & ", AZEIT Text(8) "
    sSQL = sSQL & ", KUNDNR LONG "
    sSQL = sSQL & ", FILIALE BYTE "
    sSQL = sSQL & ", KASNUM BYTE "
    sSQL = sSQL & ", LINR long"
    sSQL = sSQL & ", LPZ INTEGER"
    sSQL = sSQL & ", AGN Long "
    sSQL = sSQL & ", EAN Text(13)"
    sSQL = sSQL & ", MWST Text(1)"
    sSQL = sSQL & ", EKPR SINGLE "
    sSQL = sSQL & ", VKPR SINGLE "
    sSQL = sSQL & ", MOPREIS SINGLE "
    sSQL = sSQL & ", BELEGNR INTEGER "
    sSQL = sSQL & ", BEST1 INTEGER "
    sSQL = sSQL & ", RABKENN Text(1)"
    sSQL = sSQL & ", KK_ART Text(2)"
    sSQL = sSQL & ", BEDIENER integer "
    sSQL = sSQL & ", UMS_OK Text(1)"
    sSQL = sSQL & ", ZBONNR integer"
    sSQL = sSQL & ", ABOK BIT"
    
    sSQL = sSQL & ", dewasArtnr double "
    sSQL = sSQL & ", dewasMWST Text(6)"
    
    sSQL = sSQL & ")"
    gdBase.Execute sSQL, dbFailOnError

    
    
    
    sSQL = "Insert into Kassjour_Dewas Select "
    
    
    sSQL = sSQL & " 0 as  ARTNR  "
    sSQL = sSQL & ", cartbezlang as BEZEICH "
    sSQL = sSQL & ", dmenge as MENGE "
    sSQL = sSQL & ", dumsatz as PREIS "
    sSQL = sSQL & ", dtdatum as ADATE "
    sSQL = sSQL & ", '00:00:00' as AZEIT "
    sSQL = sSQL & ", 0 as KUNDNR "
    sSQL = sSQL & ", 0 as FILIALE "
    sSQL = sSQL & ", sikassnr as KASNUM "
    sSQL = sSQL & ", 0 as LINR "
    sSQL = sSQL & ", 0 as LPZ "
    sSQL = sSQL & ", lwgr as AGN "
    sSQL = sSQL & ", dplunr as EAN"
    sSQL = sSQL & ", '' as MWST "
    sSQL = sSQL & ", siprzmwst as dewasMWST "
    sSQL = sSQL & ", dekpreis as EKPR  "
    sSQL = sSQL & ", dstpreis as VKPR  "
    sSQL = sSQL & ", 0 as MOPREIS "
    sSQL = sSQL & ", lbonnr as BELEGNR "
    sSQL = sSQL & ", 0 as BEST1  "
    sSQL = sSQL & ", '' as RABKENN "
    sSQL = sSQL & ", 'BA' as KK_ART "
    sSQL = sSQL & ", 1 as BEDIENER  "
    sSQL = sSQL & ", 'J' as UMS_OK "
    sSQL = sSQL & ", 0 as ZBONNR "
    sSQL = sSQL & ", 1 as ABOK "
    
    sSQL = sSQL & ", dartnr as dewasArtnr  "
    

    sSQL = sSQL & " from [;DATABASE=" & cPfad & "\safe.MDB].dbtzba"
    sSQL = sSQL & " where dtdatum >= " & CLng(DateValue(sStartdatum))
    sSQL = sSQL & " and siabteilung <  20 "
'    sSQL = sSQL & " and cMarktnr = '89700053'  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
'    sSQL = "Insert into Kassjour_Dewas Select "
'
'
'    sSQL = sSQL & " 0 as  ARTNR  "
'    sSQL = sSQL & ", cartbezlang as BEZEICH "
'    sSQL = sSQL & ", dmenge as MENGE "
'    sSQL = sSQL & ", dumsatz as PREIS "
'    sSQL = sSQL & ", dtdatum as ADATE "
'    sSQL = sSQL & ", '00:00:00' as AZEIT "
'    sSQL = sSQL & ", 0 as KUNDNR "
'    sSQL = sSQL & ", 2 as FILIALE "
'    sSQL = sSQL & ", sikassnr as KASNUM "
'    sSQL = sSQL & ", 0 as LINR "
'    sSQL = sSQL & ", 0 as LPZ "
'    sSQL = sSQL & ", lwgr as AGN "
'    sSQL = sSQL & ", dplunr as EAN"
'    sSQL = sSQL & ", '' as MWST "
'    sSQL = sSQL & ", siprzmwst as dewasMWST "
'    sSQL = sSQL & ", dekpreis as EKPR  "
'    sSQL = sSQL & ", dstpreis as VKPR  "
'    sSQL = sSQL & ", 0 as MOPREIS "
'    sSQL = sSQL & ", lbonnr as BELEGNR "
'    sSQL = sSQL & ", 0 as BEST1  "
'    sSQL = sSQL & ", '' as RABKENN "
'    sSQL = sSQL & ", 'BA' as KK_ART "
'    sSQL = sSQL & ", 1 as BEDIENER  "
'    sSQL = sSQL & ", 'J' as UMS_OK "
'    sSQL = sSQL & ", 0 as ZBONNR "
'    sSQL = sSQL & ", 1 as ABOK "
'
'    sSQL = sSQL & ", dartnr as dewasArtnr  "
'
'
'    sSQL = sSQL & " from [;DATABASE=" & cPfad & "\safe.MDB].dbtzba"
'    sSQL = sSQL & " where dtdatum >= " & CLng(DateValue(sStartdatum))
'    sSQL = sSQL & " and siabteilung <  20 "
'    sSQL = sSQL & " and cMarktnr = '89700277'  "
'    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'Ende Kassjour_Dewas erstellen
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel wird erstellt...": lblx.Refresh
    
    'Artikel erstellen
    
    loeschNEW "ARTIKEL_DEWAS", gdBase

    sSQL = "Create Table ARTIKEL_DEWAS"
    sSQL = sSQL & " ( "
    sSQL = sSQL & " ARTNR double "
'    sSQL = sSQL & ", Bestand long "
    sSQL = sSQL & ", BEZEICH Text(35) "
    sSQL = sSQL & ", MWST Text(1) "
    sSQL = sSQL & ", VKPR double"
    sSQL = sSQL & ", KVKPR1 double"
    sSQL = sSQL & ", INHALT double"
    sSQL = sSQL & ", INHALTBEZ Text(3)"
    sSQL = sSQL & ", INHALTKOM Text(10)"
    sSQL = sSQL & ", AGN long "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into ARTIKEL_DEWAS Select "
    sSQL = sSQL & " dArtnr as ARTNR  "
'    sSQL = sSQL & ", dBestand as Bestand  "
    sSQL = sSQL & ", simwst as MWST  "
    sSQL = sSQL & ", left(trim(cBez),35) as Bezeich  "
    sSQL = sSQL & ", lwgr as AGN "
    sSQL = sSQL & ", cInhalt as INHALTKOM "
    sSQL = sSQL & ", dvkgrh as VKPR "
    sSQL = sSQL & ", dVK as KVKPR1 "
    sSQL = sSQL & " from [;DATABASE=" & cPfad & "\safe.MDB].dbarts"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": MwSt wird aktualisiert...": lblx.Refresh
    
    sSQL = "Update ARTIKEL_DEWAS Set MWST = 'E' where MWST = '1' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTIKEL_DEWAS Set MWST = 'V' where MWST = '2' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTIKEL_DEWAS Set MWST = 'O' where MWST = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = TimeValue(Now) & ": Inhalt wird aktualisiert...": lblx.Refresh
    
    Dim cInhaltges As String
    Dim cInhalt As String
    Dim cInhaltBez As String

    sSQL = "select * from ARTIKEL_DEWAS where INHALTKOM <> '' "
    Set rsrs = gdBase.OpenRecordset(sSQL)

    If Not rsrs.EOF Then
        rsrs.MoveFirst
    
        Do While Not rsrs.EOF
        
            cInhalt = ""
            cInhaltBez = ""
            If Not IsNull(rsrs!INHALTKOM) Then
                cInhaltges = rsrs!INHALTKOM
            End If
            
            cInhaltBez = ""
            cInhalt = ""
            If Len(cInhaltges) > 0 Then
                For i = Len(cInhaltges) To 0 Step -1
                    cInhalt = Mid(cInhaltges, 1, i)
                    If IsNumeric(cInhalt) Then
                        cInhaltBez = Right(cInhaltges, Len(cInhaltges) - i)
                        cInhaltBez = SwapStr(cInhaltBez, ".", "")
                        Exit For
                    End If
                Next i
            End If
            rsrs.Edit
            
            cInhaltBez = UCase(Left(cInhaltBez, 3))
            If InStr(1, cInhaltBez, " ") Then
                cInhaltBez = Trim(Mid(cInhaltBez, 1, InStr(1, cInhaltBez, " ")))
            End If
            
            rsrs!INHALTBEZ = UCase(Left(cInhaltBez, 3))
                    
            If cInhalt <> "" Then
                rsrs!INHALT = cInhalt
            End If
            rsrs.Update
            
            rsrs.MoveNext
        Loop

    End If

    rsrs.Close
    
    
    lblx.Caption = TimeValue(Now) & ": Bezeichnung wird aktualisiert...": lblx.Refresh
    
'    sSQL = "Update ARTIKEL set Bezeich = Replace(bezeich,' ','')"
'    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "select * from ARTIKEL_DEWAS where BEZEICH like '*" & Chr(34) & "*' "
    Set rsrs = gdBase.OpenRecordset(sSQL)

    If Not rsrs.EOF Then
        rsrs.MoveFirst
    
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEZEICH) Then
                cBez = rsrs!BEZEICH
                cBez = SwapStr(cBez, Chr(34), "")
            End If
                    
            rsrs.Edit
            rsrs!BEZEICH = cBez
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    'Ende Artikel
    
    
    'Artikel bauen
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel wird zusammengestellt...": lblx.Refresh
    
    sSQL = "Delete from ARTIKEL "
    gdBase.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("ARTIKEL", "dewasArtnr", gdBase) = False Then
        sSQL = "Alter table ARTIKEL add column dewasArtnr double"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Create Index dewasartnr on artikel (dewasartnr)"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Delete from ARTLIEF "
    gdBase.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("ARTLIEF", "dewasArtnr", gdBase) = False Then
        sSQL = "Alter table ARTLIEF add column dewasArtnr double"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Create Index dewasartnr on ARTLIEF (dewasartnr)"
        gdBase.Execute sSQL, dbFailOnError
        
        
    End If
    
    sSQL = "Insert into Artikel select artnr as dewasartnr "
    sSQL = sSQL & ", MWST  "
'    sSQL = sSQL & ", Bestand  "
    sSQL = sSQL & ", Bezeich  "
    sSQL = sSQL & ", AGN "
    sSQL = sSQL & ", Inhalt "
    sSQL = sSQL & ", Inhaltbez "
    sSQL = sSQL & ", VKPR "
    sSQL = sSQL & ", KVKPR1 "
    sSQL = sSQL & " from ARTIKEL_DEWAS "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Grundpreis wird aktualisiert...": lblx.Refresh
    
    sSQL = "Update ARTIKEL Set INHALT = 0 where Inhalt is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTIKEL Set GRUNDPREIS = 'J' where Inhalt <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": neue Artikelnummer wird vergeben...": lblx.Refresh
    
    sSQL = "Alter table ARTIKEL add column LFNR autoincrement"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTIKEL Set ARTNR = 600000 + lfnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Alter table ARTIKEL drop LFNR "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour Artnr anpassen...": lblx.Refresh
    
    sSQL = "Create Index dewasartnr on Kassjour_dewas (dewasartnr)"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Create Index artnr on Kassjour_dewas (artnr)"
    gdBase.Execute sSQL, dbFailOnError



    sSQL = "Create Index dewasMWST on Kassjour_dewas (dewasMWST)"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update Kassjour_dewas k inner join artikel a on k.dewasartnr = a.dewasartnr"
    sSQL = sSQL & " set k.artnr = a.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = TimeValue(Now) & ": MwSt wird aktualisiert...": lblx.Refresh
    
    sSQL = "Update Kassjour_Dewas Set MWST = 'E' where dewasMWST = '700' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour_Dewas Set MWST = 'V' where dewasMWST = '1900' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour_Dewas Set MWST = 'O' where dewasMWST = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour_Dewas Set ARTNR  = 0  where Artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour_Dewas Set ARTNR  = 999999 and Bezeich = 'ohne Zuordnung' where Artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artlief wird zusammengestellt...": lblx.Refresh
    
    sSQL = "Insert into ARTLIEF select artnr as dewasartnr "
    sSQL = sSQL & ", Linr "
    sSQL = sSQL & ", libesnr "
    sSQL = sSQL & ", VPE as MinMen "
    sSQL = sSQL & ", LEKPR "
    sSQL = sSQL & ", 'E' as SYNSTATUS "
    sSQL = sSQL & ", 0 as SPANNE "
    sSQL = sSQL & " from ARTLIEF_DEWAS "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTLIEF l inner join artikel a on l.dewasartnr = a.dewasartnr"
    sSQL = sSQL & " set l.artnr = a.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    If SpalteInTabellegefundenNEW("ARTEAN_DEWAS", "LFNR", gdBase) = False Then
        sSQL = "Alter table ARTEAN_DEWAS add column LFNR autoincrement"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    loeschNEW "ARTEAN_DEWAS_AR", gdBase
    sSQL = "Select * into ARTEAN_DEWAS_AR from ARTEAN_DEWAS "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": 1.EAN wird aktualisiert...": lblx.Refresh
    
    sSQL = "Update Artikel a inner join ARTEAN_DEWAS_AR e on a.dewasartnr = e.artnr"
    sSQL = sSQL & " set a.ean = e.ean "
    gdBase.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("ARTEAN_DEWAS_AR", "ERKANNT", gdBase) = False Then
        sSQL = "Alter table ARTEAN_DEWAS_AR add column ERKANNT Text(1)"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Update ARTEAN_DEWAS_AR e inner join Artikel a on e.ean = a.ean"
    sSQL = sSQL & " set e.erkannt = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ARTEAN_DEWAS_AR where erkannt = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = TimeValue(Now) & ": 2.EAN wird aktualisiert...": lblx.Refresh
    'jetzt EAN 2 füllen
    
    sSQL = "Update Artikel a inner join ARTEAN_DEWAS_AR e on a.dewasartnr = e.artnr"
    sSQL = sSQL & " set a.ean2 = e.ean "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEAN_DEWAS_AR e inner join Artikel a on e.ean = a.ean2"
    sSQL = sSQL & " set e.erkannt = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ARTEAN_DEWAS_AR where erkannt = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": 3.EAN wird aktualisiert...": lblx.Refresh
    'jetzt EAN 2 füllen
    
    sSQL = "Update Artikel a inner join ARTEAN_DEWAS_AR e on a.dewasartnr = e.artnr"
    sSQL = sSQL & " set a.ean3 = e.ean "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEAN_DEWAS_AR e inner join Artikel a on e.ean = a.ean3"
    sSQL = sSQL & " set e.erkannt = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ARTEAN_DEWAS_AR where erkannt = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    'den Rest in eine ARTEAN Tabelle , an der Kasse scannen führt den EAN an die 1.Stelle (1.EAN)
    
    If SpalteInTabellegefundenNEW("ARTEAN_DEWAS_AR", "KISS_ARTNR", gdBase) = False Then
        sSQL = "Alter table ARTEAN_DEWAS_AR add column KISS_ARTNR Long "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Not NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then 'das erste Mal
    
        sSQL = "Create Table ARTEAN_K"
        sSQL = sSQL & " ( "
        sSQL = sSQL & " ARTNR long "
        sSQL = sSQL & ", ean Text(13) "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Create Index EAN on ARTEAN_K (EAN)"
        gdBase.Execute sSQL, dbFailOnError
        
    End If
    
    
    sSQL = "Update ARTEAN_DEWAS_AR e inner join Artikel a on e.artnr = a.dewasartnr "
    sSQL = sSQL & " set e.KISS_ARTNR = a.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEAN_DEWAS_AR set kiss_artnr = 0 where kiss_artnr is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ARTEAN_DEWAS_AR where kiss_artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
        
    sSQL = "Insert into ARTEAN_K select kiss_artnr as artnr, ean from  ARTEAN_DEWAS_AR where kiss_artnr > 0 "
    sSQL = sSQL & " and not ARTEAN_DEWAS_AR.ean in (Select ean from ARTEAN_K) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'achtung mehrmaliges füllen verhindern
    
    sSQL = "Update Artikel "
    sSQL = sSQL & " set gefuehrt = 'J' "
    sSQL = sSQL & " , BONUS_OK = 'J' "
    sSQL = sSQL & " , RABATT_OK = 'J' "
    sSQL = sSQL & " , UMS_OK = 'J' "
    sSQL = sSQL & " , AWM = '0' "
    sSQL = sSQL & " , RKZ = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    
    'baue eine kleinste EK Tabelle ausser 0
    loeschNEW "KLEINEEK", gdBase
    
    sSQL = "Select min(LEKPR) as ek , artnr into KLEINEEK from ARTLIEF "
    sSQL = sSQL & " where LEKPR > 0 group by artnr order by min(lekpr) asc "
    gdBase.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("KLEINEEK", "LINR", gdBase) = False Then
        sSQL = "Alter table KLEINEEK add column LINR LONG "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Update KLEINEEK k inner join ARTLIEF a on k.artnr = a.artnr and k.ek = a.LEKPR"
    sSQL = sSQL & " set k.linr = a.linr  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update Artikel set gefuehrt = 'J' "
    sSQL = sSQL & " , BONUS_OK = 'J' "
    sSQL = sSQL & " , RABATT_OK = 'J' "
    sSQL = sSQL & " , UMS_OK = 'J' "
    sSQL = sSQL & " , AWM = '0' "
    sSQL = sSQL & " , RKZ = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel a "
    sSQL = sSQL & " set a.ekpr = 0 "
    gdBase.Execute sSQL, dbFailOnError
        
    
    sSQL = "Update Artikel a inner join KLEINEEK k on a.artnr = k.artnr "
    sSQL = sSQL & " set a.ekpr = k.ek "
    sSQL = sSQL & " , a.linr = k.linr "
    gdBase.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("ARTIKEL", "dewasArtnr", gdBase) = True Then
    
        sSQL = "drop index dewasArtnr on artikel"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Alter table ARTIKEL drop column dewasArtnr "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
   
    If SpalteInTabellegefundenNEW("ARTLIEF", "dewasArtnr", gdBase) = True Then
    
        sSQL = "drop index dewasArtnr on ARTLIEF"
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "Alter table ARTLIEF drop column dewasArtnr "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("Kassjour_dewas", "dewasArtnr", gdBase) = True Then
    
        sSQL = "drop index dewasArtnr on Kassjour_dewas"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Alter table Kassjour_dewas drop column dewasArtnr "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("Kassjour_dewas", "dewasMWST", gdBase) = True Then
    
        sSQL = "drop index dewasMWST on Kassjour_dewas"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Alter table Kassjour_dewas drop column dewasMWST "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
'    sSQL = "Delete from Kassjour where adate >= " & CLng(DateValue(sStartdatum))
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from Kassjour "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Kassjour select * from Kassjour_DEWAS"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTLIEF set "
    sSQL = sSQL & "  RKZ = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean = val(ean) where ean <> '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update artean_K set ean = val(ean) where ean <> '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean2 = val(ean2) where ean2 <> '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean3 = val(ean3) where ean3 <> '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LISRT set kuerzel = Ucase(left(liefbez,5))"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean = '' where len(ean) < 5 "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "KLEINEEK", gdBase
    loeschNEW "ARTEAN_DEWAS", gdBase
    loeschNEW "ARTLIEF_DEWAS", gdBase
    loeschNEW "Kassjour_Dewas", gdBase
    loeschNEW "ARTEAN_DEWAS_AR", gdBase
    loeschNEW "ARTIKEL_DEWAS", gdBase
        
    
    lblx.Caption = TimeValue(Now) & ": Der Dewas-Import ist fertig!": lblx.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DewasImport"
        Fehler.gsFehlertext = "Im Programmteil Dewas Import ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
    
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    anzeige "normal", "", Label1(4)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Dewas Import ist ein Fehler aufgetreten."
    
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

Private Sub txtStatus_Change()
    On Error GoTo LOKAL_ERROR
    
    Dim nProz As Long
  
    nProz = Val(txtStatus.Text)
    ShowProgress picprogress, nProz, 0, 100, True
    picprogress.Refresh

Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtstatus_Change"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


