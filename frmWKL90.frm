VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL90 
   Caption         =   "Webshop - Schnittstellen"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   Icon            =   "frmWKL90.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10080
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   240
      MaxLength       =   200
      TabIndex        =   9
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   240
      MaxLength       =   200
      TabIndex        =   7
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   240
      MaxLength       =   200
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   240
      MaxLength       =   200
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   0
      Top             =   6360
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   7
      Left            =   7800
      TabIndex        =   2
      Top             =   5880
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Script erstellen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "bestand.php: Dieses Script aktualisiert sofort bei jeder Bestansveränderung in der Warenwirtschaft den Bestand in Ihrem Webshop."
      Height          =   1575
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   9615
   End
   Begin VB.Label lblAnzeige 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "MySQL - Datenbankserver"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "MySQL - Datenbankpasswort"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "MySQL - Datenbankbenutzer"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "MySQL - Datenbankname"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Webshop - Schnittstellen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmWKL90"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Select Case Index
    Case 7
    
        If Text2(0).Text <> "" And Text2(1).Text <> "" And Text2(2).Text <> "" And Text2(3).Text <> "" Then
            schreibe_Php_Script_Bestand Text2(0).Text, Text2(1).Text, Text2(2).Text, Text2(3).Text
        Else
            anzeige "normal", "Bitte ALLE Eingabefelder ausfüllen!", lblAnzeige
        End If
    Case 1
        Unload frmWKL90
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Webshop - Schnittstellen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub schreibe_Php_Script_Bestand(sServer As String, sDabaname As String, sDabaUser As String, sDabapass As String)
On Error GoTo LOKAL_ERROR
    
    Dim cSatz           As String
    Dim cPfad           As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cDatname        As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "BOX\"
    cPfad = UCase$(cPfad)

    cDatname = "bestand.php"
    Kill cPfad & cDatname
    
    iFileNr = FreeFile
    Open cPfad & cDatname For Binary As #iFileNr
    
    
    
    
    
    
    cSatz = "<?php" & vbCrLf
    cSatz = cSatz & "error_reporting(E_PARSE);" & vbCrLf

    cSatz = cSatz & "// ################################# URL ###################################################################################" & vbCrLf
    cSatz = cSatz & "// http://www.officeprosystem.de/test/kiss.php?quelle=kiss&tab=artikel&spalteb=bestand&spaltea=artnr&artnr=555555&menge=5656" & vbCrLf

    cSatz = cSatz & "$quelle=$_GET[" & Chr(34) & "quelle" & Chr(34) & "];" & vbCrLf

    cSatz = cSatz & "if($quelle==" & Chr(34) & "kiss" & Chr(34) & "){" & vbCrLf
    cSatz = cSatz & "$server = " & Chr(34) & "" & sServer & "" & Chr(34) & ";" & vbCrLf
    cSatz = cSatz & "$datenbank=" & Chr(34) & "" & sDabaname & "" & Chr(34) & ";" & vbCrLf
    cSatz = cSatz & "$benutzer=" & Chr(34) & "" & sDabaUser & "" & Chr(34) & ";" & vbCrLf
    cSatz = cSatz & "$passwort=" & Chr(34) & "" & sDabapass & "" & Chr(34) & ";" & vbCrLf
    cSatz = cSatz & "$tabelle=$_GET[" & Chr(34) & "tab" & Chr(34) & "];" & vbCrLf
    cSatz = cSatz & "$spalteb=$_GET[" & Chr(34) & "spalteb" & Chr(34) & "];" & vbCrLf
    cSatz = cSatz & "$spaltea=$_GET[" & Chr(34) & "spaltea" & Chr(34) & "];" & vbCrLf
    cSatz = cSatz & "$artnr=$_GET[" & Chr(34) & "artnr" & Chr(34) & "];" & vbCrLf
    cSatz = cSatz & "$menge=$_GET[" & Chr(34) & "menge" & Chr(34) & "];" & vbCrLf

    cSatz = cSatz & "// Verbindung erstellen" & vbCrLf
    cSatz = cSatz & "$link = mysql_connect($server, $benutzer, $passwort);" & vbCrLf
    cSatz = cSatz & "if (! $link)" & vbCrLf
    cSatz = cSatz & "die(" & Chr(34) & "Verbindung gescheitert" & Chr(34) & ");" & vbCrLf
    cSatz = cSatz & "mysql_select_db($datenbank);" & vbCrLf

    cSatz = cSatz & "$suche=" & Chr(34) & "select $spalteb from $tabelle where $spaltea = '$artnr'" & Chr(34) & ";" & vbCrLf
    cSatz = cSatz & "$treffer=mysql_query($suche) or die (" & Chr(34) & "Fehlermeldung=" & Chr(34) & ".mysql_error());" & vbCrLf
    cSatz = cSatz & "$treffer=mysql_fetch_row($treffer);" & vbCrLf
    cSatz = cSatz & "if($treffer){" & vbCrLf
    cSatz = cSatz & "$edit=" & Chr(34) & "update $tabelle set $spalteb = $menge where $spaltea = '$artnr'" & Chr(34) & ";" & vbCrLf
    cSatz = cSatz & "mysql_query($edit) or die (" & Chr(34) & "Fehlermeldung=" & Chr(34) & ".mysql_error());" & vbCrLf
    cSatz = cSatz & "echo $menge;" & vbCrLf
    cSatz = cSatz & "}" & vbCrLf
    cSatz = cSatz & "else{" & vbCrLf
    cSatz = cSatz & "echo 'Artikel nicht gefunden';" & vbCrLf
    cSatz = cSatz & "}" & vbCrLf

    cSatz = cSatz & "}" & vbCrLf
    cSatz = cSatz & "else{" & vbCrLf

    cSatz = cSatz & "echo " & Chr(34) & "nö" & Chr(34) & ";" & vbCrLf
    cSatz = cSatz & "}" & vbCrLf
    cSatz = cSatz & "?>"
    

    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    Close iFileNr
    
    MsgBox "Die Datei befindet sich hier: " & cPfad & "BOX\bestand.php", vbInformation, "Winkiss Hinweis:"
    
    Screen.MousePointer = 0
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "schreibe_Php_Script_Bestand"
        Fehler.gsFehlertext = "Im Programmteil Webshop - Schnittstellen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    alternativFarbform Me, lblUeberschrift
    Modul6.Skalieren Me, True, True: Schrift Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Webshop - Schnittstellen ist ein Fehler aufgetreten."
    
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





