Attribute VB_Name = "Modul6"
Option Explicit


'die muss bleiben dez 03
Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, ByVal lParam As Long) _
        As Long

'//Thomas
Private Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRoothPath As String, ByVal lpInputName As String, ByVal lpOutputName As String) As Long
Const MAX_PATH = 260
Public Function ermfak(cArtNr As String) As Long
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermfak = 0
    
    If IsNumeric(cArtNr) Then
        sSQL = "Select Faktor from zuordean where artnr = " & cArtNr
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Faktor) Then
                ermfak = Val(rsrs!Faktor)
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "ermfak"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgpEAN1(cArtNr As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermgpEAN1 = ""
    
    If IsNumeric(cArtNr) Then
        
        sSQL = "Select GPEAN from zuordean where artnr = " & cArtNr
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!GPEAN) Then
                ermgpEAN1 = Trim(rsrs!GPEAN)
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "ermgpEAN1"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermKUBMENGE(cArtNr As String) As Long
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermKUBMENGE = 0
    
    If IsNumeric(cArtNr) Then
        
        sSQL = "Select sum(Bestelltmenge) as maxi from KUNDBEST where artnr = " & cArtNr
        sSQL = sSQL & " and StatusARTIKEL = 'INBESTELLUNG' "
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!maxi) Then
                ermKUBMENGE = Val(rsrs!maxi)
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "ermKUBMENGE"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermKUBMENGEforGeliefert(cArtNr As String) As Long
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermKUBMENGEforGeliefert = 0
    
    If IsNumeric(cArtNr) Then
        
        sSQL = "Select sum(Bestelltmenge) as maxi from KUNDBEST where artnr = " & cArtNr
        sSQL = sSQL & " and (StatusARTIKEL = 'INBESTELLUNG' or StatusARTIKEL = 'BESTELLT')"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!maxi) Then
                ermKUBMENGEforGeliefert = Val(rsrs!maxi)
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "ermKUBMENGEforGeliefert"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Sub Schrift(Frm As Form)
    On Error GoTo LOKAL_ERROR

    Dim i           As Integer
    Dim FontFak     As Integer
    Dim sName As String
    
    If Val(gsFontsize) = 0 Then
        gsFontsize = 12
    End If
    
    If gsFontsize = 0 Then
        gsFontsize = 12
    End If
    
'    If IsNumeric(gsFontsize) = False Then
'        gsFontsize = "12"
'    End If
    
    If Trim(gsFont) = "" Then
        gsFont = "Arial"
    End If
    
    FontFak = 12 - gsFontsize
            
    For i = 0 To Frm.Controls.Count - 1
    
        sName = Frm.Controls(i).name

        If TypeOf Frm.Controls(i) Is MSFlexGrid Then 'alle Frames
        
            Frm.Controls(i).Font = gsFont
            Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
        End If
        
        If TypeOf Frm.Controls(i) Is TabStrip Then 'alle Frames
            Frm.Controls(i).Font = gsFont
        End If
        
        If TypeOf Frm.Controls(i) Is Frame Then 'alle Frames
            Frm.Controls(i).Font = gsFont
            Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
            
        End If
        
        If TypeOf Frm.Controls(i) Is TextBox Then 'alle Frames
            Frm.Controls(i).Font = gsFont
            Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
            
        End If
        
        If TypeOf Frm.Controls(i) Is Label Then 'alle Labels
            If Frm.Controls(i).Tag = "FESTEBREITE" Then
                
                Frm.Controls(i).Font = "Courier New"
                Frm.Controls(i).FontBold = False
                Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
            
            Else
            
                Frm.Controls(i).Font = gsFont
                Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
            End If
        End If
        
        If TypeOf Frm.Controls(i) Is CheckBox Then 'alle Checkboxen
            Frm.Controls(i).Font = gsFont
            Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
        End If
        
        
        If TypeOf Frm.Controls(i) Is sevCommand3.Command Then 'alle Buttons
            Frm.Controls(i).Font = gsFont
            Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
            
        End If
        
        If TypeOf Frm.Controls(i) Is SSCommand Then 'alle Buttons
            Frm.Controls(i).Font = gsFont
            Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
            
        End If
        
        If TypeOf Frm.Controls(i) Is OptionButton Then 'alle Optionsfelder
            Frm.Controls(i).Font = gsFont
            Frm.Controls(i).FontSize = Frm.Controls(i).FontSize - FontFak
        End If
        
'        If TypeOf Frm.Controls(i) Is Calendar Then
''            sElement = "Calendar"
'            Frm.Controls(i).DayFont.name = gsFont
'            Frm.Controls(i).GridFont.name = gsFont
'            Frm.Controls(i).DayFont.Size = Frm.Controls(i).DayFont.Size - FontFak
'            Frm.Controls(i).GridFont.Size = Frm.Controls(i).GridFont.Size - FontFak
'
'            Frm.Controls(i).TitleFont.name = gsFont
'            Frm.Controls(i).TitleFont.Size = Frm.Controls(i).TitleFont.Size - FontFak
'        End If


    Next i
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "Schrift"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. " & Frm.name & " " & sName & " " & gsFontsize & " " & gsFont & " " & FontFak
        
        Fehlermeldung1
    End If
End Sub
Public Sub Log(Frm As Form)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If gbLokalModus = False Then
    
        LogtoStart Frm
        
    End If
    
    
    Exit Sub
    
    
LOKAL_ERROR:
    If err.Number = 91 Or err.Number = 3260 Or err.Number = 3192 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "Log"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Public Sub SpeicherCaption(Frm As Form)
    On Error GoTo LOKAL_ERROR


    Dim i As Integer
    Dim sSQL As String
    Dim sBez As String


    
    For i = 0 To Frm.Controls.Count - 1

        If TypeOf Frm.Controls(i) Is Frame Then 'alle Frames
            If Frm.Controls(i).Caption <> "" Then
                If Len(Frm.Controls(i).Caption) > 1 Then
                    sBez = Frm.Controls(i).Caption
                    sBez = SwapStr(sBez, "'", "")
                    sBez = SwapStr(sBez, ".", "")
                   
                    
                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH) values "
                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
            
        End If
        
        If TypeOf Frm.Controls(i) Is Label Then 'alle Labels
            If Frm.Controls(i).Caption <> "" Then
                If Len(Frm.Controls(i).Caption) > 1 Then
                
                    sBez = Frm.Controls(i).Caption
                    sBez = SwapStr(sBez, "'", "")
                    sBez = SwapStr(sBez, ".", "")
                    
                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH ) values "
                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
        End If
        
        If TypeOf Frm.Controls(i) Is CheckBox Then 'alle Checkboxen
            If Frm.Controls(i).Caption <> "" Then
                If Len(Frm.Controls(i).Caption) > 1 Then
                
                    sBez = Frm.Controls(i).Caption
                    sBez = SwapStr(sBez, "'", "")
                    sBez = SwapStr(sBez, ".", "")
                    
                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH ) values "
                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
        End If
        
        If TypeOf Frm.Controls(i) Is sevCommand3.Command Then 'alle Buttons
            If Frm.Controls(i).Caption <> "" Then
                If Len(Frm.Controls(i).Caption) > 1 Then

                    sBez = Frm.Controls(i).Caption
                    sBez = SwapStr(sBez, "'", "")
                    sBez = SwapStr(sBez, ".", "")
                    
                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH ) values "
                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
            
        End If
        
        If TypeOf Frm.Controls(i) Is SSCommand Then 'alle Buttons
            If Frm.Controls(i).Caption <> "" Then
                If Len(Frm.Controls(i).Caption) > 1 Then
                
                    sBez = Frm.Controls(i).Caption
                    sBez = SwapStr(sBez, "'", "")
                    sBez = SwapStr(sBez, ".", "")
                    
                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH ) values "
                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
            
        End If
        
        If TypeOf Frm.Controls(i) Is OptionButton Then 'alle Optionsfelder
            If Frm.Controls(i).Caption <> "" Then
                If Len(Frm.Controls(i).Caption) > 1 Then
                
                    sBez = Frm.Controls(i).Caption
                    sBez = SwapStr(sBez, "'", "")
                    sBez = SwapStr(sBez, ".", "")
                
                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH ) values "
                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
        End If


    Next i
    


    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "SpeicherCaption"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub SpeicherBTNCaption(Frm As Form)
    On Error GoTo LOKAL_ERROR


    Dim i As Integer
    Dim sSQL As String
    Dim sBez As String
    Dim sName As String
    
    Dim sIndex As String


    
    For i = 0 To Frm.Controls.Count - 1

'        If TypeOf Frm.Controls(i) Is Frame Then 'alle Frames
'            If Frm.Controls(i).Caption <> "" Then
'                If Len(Frm.Controls(i).Caption) > 1 Then
'                    sBez = Frm.Controls(i).Caption
'                    sBez = SwapStr(sBez, "'", "")
'                    sBez = SwapStr(sBez, ".", "")
'
'
'                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH) values "
'                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
'                    gdBase.Execute sSQL, dbFailOnError
'                End If
'            End If
'
'        End If
'
'        If TypeOf Frm.Controls(i) Is Label Then 'alle Labels
'            If Frm.Controls(i).Caption <> "" Then
'                If Len(Frm.Controls(i).Caption) > 1 Then
'
'                    sBez = Frm.Controls(i).Caption
'                    sBez = SwapStr(sBez, "'", "")
'                    sBez = SwapStr(sBez, ".", "")
'
'                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH ) values "
'                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
'                    gdBase.Execute sSQL, dbFailOnError
'                End If
'            End If
'        End If
        
'        If TypeOf Frm.Controls(i) Is CheckBox Then 'alle Checkboxen
'            If Frm.Controls(i).Caption <> "" Then
'                If Len(Frm.Controls(i).Caption) > 1 Then
'
'                    sBez = Frm.Controls(i).Caption
'                    sBez = SwapStr(sBez, "'", "")
'                    sBez = SwapStr(sBez, ".", "")
'
'                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH ) values "
'                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
'                    gdBase.Execute sSQL, dbFailOnError
'                End If
'            End If
'        End If







        
        If TypeOf Frm.Controls(i) Is sevCommand3.Command Then 'alle Buttons
            If Frm.Controls(i).Caption <> "" Then
                If Len(Frm.Controls(i).Caption) > 1 Then
                
                    sIndex = ""

                    sBez = Frm.Controls(i).Caption
                    sBez = SwapStr(sBez, "'", "")
                    sBez = SwapStr(sBez, ".", "")
                    
                    
                    
                    sName = Frm.Controls(i).name
                    sName = SwapStr(sName, "'", "")
                    sName = SwapStr(sName, ".", "")
                    
                    sIndex = Frm.Controls(i).index
                    
                    
                    If sIndex <> "" Then
                        sName = sName & "(" & sIndex & ")"
                    End If
                    
                    sSQL = "Insert into BTNBESCHRIFTUNG (frmName, BTNText,CMDName, NrInForm, DEUTSCH ) values "
                    sSQL = sSQL & "( '" & Frm.name & "','" & sBez & "','" & sName & "'," & i & ",'" & sBez & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
            
        End If
        
        If TypeOf Frm.Controls(i) Is SSCommand Then 'alle Buttons
            If Frm.Controls(i).Caption <> "" Then
                If Len(Frm.Controls(i).Caption) > 1 Then
                
                    sIndex = ""
                
                    sBez = Frm.Controls(i).Caption
                    sBez = SwapStr(sBez, "'", "")
                    sBez = SwapStr(sBez, ".", "")
                    
                    sName = Frm.Controls(i).name
                    sName = SwapStr(sName, "'", "")
                    sName = SwapStr(sName, ".", "")
                    
                    sIndex = Frm.Controls(i).index
                    
                    If sIndex <> "" Then
                        sName = sName & "(" & sIndex & ")"
                    End If
                    
                    sSQL = "Insert into BTNBESCHRIFTUNG (frmName, BTNText,CMDName , NrInForm, DEUTSCH ) values "
                    sSQL = sSQL & "( '" & Frm.name & "','" & sBez & "','" & sName & "'," & i & ",'" & sBez & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
            
        End If
        
'        If TypeOf Frm.Controls(i) Is OptionButton Then 'alle Optionsfelder
'            If Frm.Controls(i).Caption <> "" Then
'                If Len(Frm.Controls(i).Caption) > 1 Then
'
'                    sBez = Frm.Controls(i).Caption
'                    sBez = SwapStr(sBez, "'", "")
'                    sBez = SwapStr(sBez, ".", "")
'
'                    sSQL = "Insert into LANG (frmName , NrInForm, DEUTSCH ) values "
'                    sSQL = sSQL & "( '" & Frm.name & "'," & i & ",'" & sBez & "')"
'                    gdBase.Execute sSQL, dbFailOnError
'                End If
'            End If
'        End If


    Next i
    


    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Then
        Resume Next
    ElseIf err.Number = 343 Then
        
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "SpeicherBTNCaption"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. " & Frm.name
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Function ermCaption(formname As String, i As Integer) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    sSQL = "Select Lang1 as ergstring from LANG where frmname = '" & formname & "'"
    sSQL = sSQL & " and Nrinform = " & i
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!ergstring) Then
            ermCaption = rs!ergstring
        Else
            ermCaption = ""
        End If
    Else
        ermCaption = ""
    End If
    rs.Close: Set rs = Nothing
        
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "ermCaption"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Function
Public Sub ZeigeCaption(Frm As Form)
    On Error GoTo LOKAL_ERROR


    Dim i As Integer
    Dim sSQL As String


    
    For i = 0 To Frm.Controls.Count - 1

        If TypeOf Frm.Controls(i) Is Frame Then 'alle Frames
            If Len(Frm.Controls(i).Caption) > 1 Then Frm.Controls(i).Caption = ermCaption(Frm.name, i)
        End If
        
        If TypeOf Frm.Controls(i) Is Label Then 'alle Labels
            If Len(Frm.Controls(i).Caption) > 1 Then Frm.Controls(i).Caption = ermCaption(Frm.name, i)
        End If
        
        If TypeOf Frm.Controls(i) Is CheckBox Then 'alle Checkboxen
            If Len(Frm.Controls(i).Caption) > 1 Then Frm.Controls(i).Caption = ermCaption(Frm.name, i)
        End If
        
        If TypeOf Frm.Controls(i) Is sevCommand3.Command Then 'alle Buttons
            If Len(Frm.Controls(i).Caption) > 1 Then Frm.Controls(i).Caption = ermCaption(Frm.name, i)
        End If
        
        If TypeOf Frm.Controls(i) Is SSCommand Then 'alle Buttons
            If Len(Frm.Controls(i).Caption) > 1 Then Frm.Controls(i).Caption = ermCaption(Frm.name, i)
        End If
        
        If TypeOf Frm.Controls(i) Is OptionButton Then 'alle Optionsfelder
            If Len(Frm.Controls(i).Caption) > 1 Then Frm.Controls(i).Caption = ermCaption(Frm.name, i)
        End If


    Next i
    

    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "ZeigeCaption"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Public Sub Skalieren(Frm As Form, ScaleForm As Boolean, ScaleFont As Boolean)
    
    Dim i As Integer
    Dim SFX As Single
    Dim SFY As Single
    
    Frm.BorderStyle = 2
    Frm.WindowState = 0
    
    
    Const dWidth = 12000 'Screen.Width zur Designzeit
    Const DHeight = 9000 'Screen.Height zur Designzeit

    If FileExists(App.Path & "\NoSkalieren.cfg") Then
        SFX = 1
        SFY = 1
    Else
        SFX = Screen.Width / dWidth 'Skalierfaktor X-Achse
        SFY = Screen.Height / DHeight 'Skalierfaktor Y_Achse
        
    End If

    On Error Resume Next
    If ScaleForm Then
        Frm.Width = Frm.Width * SFX
        Frm.Height = Frm.Height * SFY
        Frm.Left = Frm.Left * SFX
        Frm.Top = Frm.Top * SFY
    End If

    For i = 0 To Frm.Controls.Count - 1
         If TypeOf Frm.Controls(i) Is line Then 'alle lines
             Frm.Controls(i).X1 = Frm.Controls(i).X1 * SFX
             Frm.Controls(i).X2 = Frm.Controls(i).X2 * SFX
             Frm.Controls(i).Y1 = Frm.Controls(i).Y1 * SFY
             Frm.Controls(i).Y2 = Frm.Controls(i).Y2 * SFY
         End If

         Frm.Controls(i).Left = Frm.Controls(i).Left * SFX
         Frm.Controls(i).Top = Frm.Controls(i).Top * SFY
         Frm.Controls(i).Width = Frm.Controls(i).Width * SFX
         Frm.Controls(i).Height = Frm.Controls(i).Height * SFY
         
         If ScaleFont Then
             Frm.Controls(i).Font.Size = Frm.Controls(i).Font.Size * SFY
             Frm.Controls(i).HeadFont.Size = Frm.Controls(i).HeadFont.Size * SFY
         End If
    Next i
    
    On Error GoTo 0
End Sub
Public Sub Skalieren_Kasse(Frm As Form, ScaleForm As Boolean, ScaleFont As Boolean)
    
    Dim i As Integer
    Dim SFX As Single
    Dim SFY As Single
    
    Frm.BorderStyle = 2
    Frm.WindowState = 0
    
    
    Const dWidth = 12000 'Screen.Width zur Designzeit
    Const DHeight = 9000 'Screen.Height zur Designzeit

    If FileExists(App.Path & "\NoSkalieren.cfg") Then
        SFX = 1
        SFY = 1
    Else
        SFX = Screen.Width / dWidth 'Skalierfaktor X-Achse
        SFY = Screen.Height / DHeight 'Skalierfaktor Y_Achse
        
    End If

    On Error Resume Next
    If ScaleForm Then
        Frm.Width = Frm.Width * SFX
        Frm.Height = Frm.Height * SFY
        Frm.Left = Frm.Left * SFX
        Frm.Top = Frm.Top * SFY
    End If

    For i = 0 To Frm.Controls.Count - 1
    
        
        If TypeOf Frm.Controls(i) Is line Then 'alle lines
            Frm.Controls(i).X1 = Frm.Controls(i).X1 * SFX
            Frm.Controls(i).X2 = Frm.Controls(i).X2 * SFX
            Frm.Controls(i).Y1 = Frm.Controls(i).Y1 * SFY
            Frm.Controls(i).Y2 = Frm.Controls(i).Y2 * SFY
        End If
        
        If TypeOf Frm.Controls(i) Is Frame Then 'alle frames
        
            If Frm.Controls(i).name = "Frame40" Then

                Frm.Controls(i).Top = Frm.Controls(i).Top * SFY
                
            Else
                Frm.Controls(i).Left = Frm.Controls(i).Left * SFX
                Frm.Controls(i).Top = Frm.Controls(i).Top * SFY
                Frm.Controls(i).Width = Frm.Controls(i).Width * SFX
                Frm.Controls(i).Height = Frm.Controls(i).Height * SFY
            End If
            

         
        ElseIf TypeOf Frm.Controls(i) Is Command Then 'alle Commands
         
            If Frm.Controls(i).ToolTipTitle = "Visa" Or Frm.Controls(i).ToolTipTitle = "Diners Club" Or Frm.Controls(i).ToolTipTitle = "Eurocard / Mastercard" _
            Or Frm.Controls(i).ToolTipTitle = "American Express" Or Frm.Controls(i).ToolTipTitle = "EC-Karte" Or Frm.Controls(i).ToolTipTitle = "Sonstige" _
            Or Frm.Controls(i).ToolTipTitle = "Alipay" Or Frm.Controls(i).ToolTipTitle = "Applepay" Or Frm.Controls(i).ToolTipTitle = "Googlepay" Or Frm.Controls(i).ToolTipTitle = "Paypal" Then
                
            Else
                Frm.Controls(i).Left = Frm.Controls(i).Left * SFX
                Frm.Controls(i).Top = Frm.Controls(i).Top * SFY
                Frm.Controls(i).Width = Frm.Controls(i).Width * SFX
                Frm.Controls(i).Height = Frm.Controls(i).Height * SFY
            End If
        ElseIf TypeOf Frm.Controls(i) Is SSCommand Then 'alle Commands
         
            If Frm.Controls(i).ToolTipText = "1 Cent" _
            Or Frm.Controls(i).ToolTipText = "2 Cent" _
            Or Frm.Controls(i).ToolTipText = "5 Cent" _
            Or Frm.Controls(i).ToolTipText = "10 Cent" _
            Or Frm.Controls(i).ToolTipText = "20 Cent" _
            Or Frm.Controls(i).ToolTipText = "50 Cent" _
            Or Frm.Controls(i).ToolTipText = "1 Euro" _
            Or Frm.Controls(i).ToolTipText = "2 Euro" _
            Or Frm.Controls(i).ToolTipText = "5 Euro" _
            Or Frm.Controls(i).ToolTipText = "10 Euro" _
            Or Frm.Controls(i).ToolTipText = "20 Euro" _
            Or Frm.Controls(i).ToolTipText = "50 Euro" _
            Or Frm.Controls(i).ToolTipText = "100 Euro" _
            Or Frm.Controls(i).ToolTipText = "200 Euro" _
            Or Frm.Controls(i).ToolTipText = "500 Euro" _
            Then
                
            Else
                Frm.Controls(i).Left = Frm.Controls(i).Left * SFX
                Frm.Controls(i).Top = Frm.Controls(i).Top * SFY
                Frm.Controls(i).Width = Frm.Controls(i).Width * SFX
                Frm.Controls(i).Height = Frm.Controls(i).Height * SFY
            End If
        Else
        
            Frm.Controls(i).Left = Frm.Controls(i).Left * SFX
            Frm.Controls(i).Top = Frm.Controls(i).Top * SFY
            Frm.Controls(i).Width = Frm.Controls(i).Width * SFX
            Frm.Controls(i).Height = Frm.Controls(i).Height * SFY
        End If

        If ScaleFont Then
            Frm.Controls(i).Font.Size = Frm.Controls(i).Font.Size * SFY
            Frm.Controls(i).HeadFont.Size = Frm.Controls(i).HeadFont.Size * SFY
        End If
    Next i
    
    
    
    On Error GoTo 0
End Sub
Public Sub allObjectsAusschalten(Frm As Form)
    On Error GoTo LOKAL_ERROR
    Dim i As Integer

    For i = 0 To Frm.Controls.Count - 1
        If TypeOf Frm.Controls(i) Is Frame Then 'alle Frames
            Frm.Controls(i).Enabled = False
        End If
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "allObjectsAusschalten"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub allObjectsEinschalten(Frm As Form)
    On Error GoTo LOKAL_ERROR
    Dim i As Integer

    For i = 0 To Frm.Controls.Count - 1
        If TypeOf Frm.Controls(i) Is Frame Then 'alle Frames
            Frm.Controls(i).Enabled = True
        End If
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "allObjectsEinschalten"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub alternativFarbform(Frm As Form, sUberschrift As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim cPfad   As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    
    If gbLokalModus = True Then
        If gbLocalSec Then
            If gbAutoLokalModus = False Then
                glH2 = vbRed
            End If
        End If
    End If
    
    Frm.BackColor = glH2
    Frm.Icon = frmWKL00.Icon
    For i = 0 To Frm.Controls.Count - 1
        
        If TypeOf Frm.Controls(i) Is Frame Then 'alle Frames
            Frm.Controls(i).BackColor = glH2
            Frm.Controls(i).ForeColor = glS1
        End If
        
        If TypeOf Frm.Controls(i) Is Label Then 'alle Labels
            If Frm.Controls(i).Tag = "Shape" Then
            
            Else
                Frm.Controls(i).BackColor = glH2
                Frm.Controls(i).ForeColor = glS1
            End If
        End If
        
        If TypeOf Frm.Controls(i) Is line Then 'alle Linien
            Frm.Controls(i).BorderColor = glU1
        End If
        
        If TypeOf Frm.Controls(i) Is CheckBox Then 'alle Checkboxen
            Frm.Controls(i).BackColor = glH2
            Frm.Controls(i).ForeColor = glS1
        End If
        
        
        If TypeOf Frm.Controls(i) Is OptionButton Then 'alle Optionsfelder
            Frm.Controls(i).BackColor = glH2
            Frm.Controls(i).ForeColor = glS1
        End If
        
        If TypeOf Frm.Controls(i) Is SSCommand Then 'alle sscommands
            Frm.Controls(i).Outline = False
        End If
        
        If TypeOf Frm.Controls(i) Is PictureBox Then 'alle PictureBoxen
            Frm.Controls(i).BackColor = glH2
        End If
        
'        If TypeOf Frm.Controls(i) Is Command Then 'alle Commands
'            Frm.Controls(i).BackColorFrom = glButtonHintergrund_from
'            Frm.Controls(i).BackColorTo = glButtonHintergrund_to
'            Frm.Controls(i).HoverColorFrom = glButtonMouseMove_Hintergrund_from
'            Frm.Controls(i).HoverColorTo = glButtonMouseMove_Hintergrund_to
'
'            Frm.Controls(i).BorderColorHover = glButtonMouseMove_Bordercolor
'            Frm.Controls(i).BorderColor = glButtonBordercolor
'
'            Frm.Controls(i).ForeColorHover = glButtonMouseMove_Forecolor
'            Frm.Controls(i).ForeColor = glButtonForecolor
'
'        End If
        
        If TypeOf Frm.Controls(i) Is Command Then 'alle Commands
            Frm.Controls(i).BackColorFrom = glButtonHintergrund_from
            Frm.Controls(i).BackColorTo = glButtonHintergrund_to
            Frm.Controls(i).HoverColorFrom = glButtonMouseMove_Hintergrund_from
            Frm.Controls(i).HoverColorTo = glButtonMouseMove_Hintergrund_to
            
            Frm.Controls(i).BorderColorHover = glButtonMouseMove_Bordercolor
            Frm.Controls(i).BorderColor = glButtonBordercolor
            
            Frm.Controls(i).ForeColorHover = glButtonMouseMove_Forecolor
            Frm.Controls(i).ForeColor = glButtonForecolor
            
            If Frm.Controls(i).ToolTipTitle = "Spaltenanordung" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Tabelle.jpg")
            End If
            If Frm.Controls(i).ToolTipTitle = "Kalender" Then
                
                
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Kalender.jpg")
                Frm.Controls(i).BackColorFrom = vbWhite
                Frm.Controls(i).BackColorTo = vbWhite
                Frm.Controls(i).PictureAlign = 3
            End If
            
            
            
            If Frm.Controls(i).ToolTipTitle = "Tastatur" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Tastatur.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Vor" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Vor.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Zurück" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Zurück.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Visa" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Visa.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Diners Club" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Diners-Club.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "American Express" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "American-Express.jpg")
            End If
            
            
            If Frm.Controls(i).ToolTipTitle = "Eurocard / Mastercard" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Mastercard.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "EC-Karte" Then
                If gsECBILD = "1" Then
                    Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "EC.jpg")
                Else
                    Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Maestro.jpg")
                End If
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Sonstige" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "diverse.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Wechseln" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "switch.jpg")
            End If
            
        End If
        

    Next i
    
    sUberschrift.ForeColor = glU1 'Überschrift
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "alternativFarbform"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
       
    End If
End Sub
Public Sub Farbform(Frm As Form, sUberschrift As Label)
    On Error GoTo LOKAL_ERROR

    Dim i       As Integer
    Dim cPfad   As String
    
    
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If gbLokalModus = True Then
        If gbLocalSec Then
            If gbAutoLokalModus = False Then
                glH1 = vbGreen
            End If
        End If
    End If
    
    If gbnachkomp = True Then
        glH1 = vbWhite
    End If
    
    Frm.BackColor = glH1    'Hintergrund H1 für das Formular
    Frm.Icon = frmWKL00.Icon
    
    Frm.Caption = gsPname & " " & Frm.Caption
    
    For i = 0 To Frm.Controls.Count - 1
        
        If TypeOf Frm.Controls(i) Is Frame Then 'alle Frames
            Frm.Controls(i).BackColor = glH1
            Frm.Controls(i).ForeColor = glS1
        End If
        
        If TypeOf Frm.Controls(i) Is Label Then 'alle Labels
            If Frm.Controls(i).Tag = "Shape" Then
            
            Else
                Frm.Controls(i).BackColor = glH1
                Frm.Controls(i).ForeColor = glS1
            End If
        End If
        
        If TypeOf Frm.Controls(i) Is CheckBox Then 'alle Checkboxen
            Frm.Controls(i).BackColor = glH1
            Frm.Controls(i).ForeColor = glS1
        End If
        
        If TypeOf Frm.Controls(i) Is line Then 'alle Linien
            Frm.Controls(i).BorderColor = glU1
        End If
        
        If TypeOf Frm.Controls(i) Is SSCommand Then 'alle sscommands
            Frm.Controls(i).Outline = False
        End If
        
        If TypeOf Frm.Controls(i) Is Command Then 'alle Commands
            Frm.Controls(i).BackColorFrom = glButtonHintergrund_from
            Frm.Controls(i).BackColorTo = glButtonHintergrund_to
            Frm.Controls(i).HoverColorFrom = glButtonMouseMove_Hintergrund_from
            Frm.Controls(i).HoverColorTo = glButtonMouseMove_Hintergrund_to
            
            Frm.Controls(i).BorderColorHover = glButtonMouseMove_Bordercolor
            Frm.Controls(i).BorderColor = glButtonBordercolor
            
            Frm.Controls(i).ForeColorHover = glButtonMouseMove_Forecolor
            Frm.Controls(i).ForeColor = glButtonForecolor
            
            If Frm.Controls(i).ToolTipTitle = "Spaltenanordung" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Tabelle.jpg")
            End If
            If Frm.Controls(i).ToolTipTitle = "Kalender" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Kalender.jpg")
                Frm.Controls(i).BackColorFrom = vbWhite
                Frm.Controls(i).BackColorTo = vbWhite
                Frm.Controls(i).PictureAlign = 3
                
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Tastatur" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Tastatur.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Vor" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Vor.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Zurück" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Zurück.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Visa" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Visa.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Diners Club" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Diners-Club.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "American Express" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "American-Express.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Eurocard / Mastercard" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Mastercard.jpg")
            End If
            
            
            

            
            If Frm.Controls(i).ToolTipTitle = "EC-Karte" Then
                If gsECBILD = "1" Then
                    Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "EC.jpg")
                Else
                    Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Maestro.jpg")
                End If
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Sonstige" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "diverse.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Wechseln" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "switch.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Rechts" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Rechts.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Links" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Links.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Rauf" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Rauf.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Runter" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Runter.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "SQL Server" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "sqlserver.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "futura" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "futura.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "WinEwws" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "WinEwws.gif")
            End If
            
        End If
        
        If TypeOf Frm.Controls(i) Is OptionButton Then 'alle Optionsfelder
            Frm.Controls(i).BackColor = glH1
            Frm.Controls(i).ForeColor = glS1
        End If

        If TypeOf Frm.Controls(i) Is PictureBox Then 'alle PictureBoxen
            Frm.Controls(i).BackColor = glH1
        End If

    Next i
    
    sUberschrift.ForeColor = glU1 'Überschrift

    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "Farbform"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub Farbform_nurButtons(Frm As Form, sUberschrift As Label)
    On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim cPfad   As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    For i = 0 To Frm.Controls.Count - 1
        
        If TypeOf Frm.Controls(i) Is Command Then 'alle Commands
            Frm.Controls(i).BackColorFrom = glButtonHintergrund_from
            Frm.Controls(i).BackColorTo = glButtonHintergrund_to
            Frm.Controls(i).HoverColorFrom = glButtonMouseMove_Hintergrund_from
            Frm.Controls(i).HoverColorTo = glButtonMouseMove_Hintergrund_to
            
            Frm.Controls(i).BorderColorHover = glButtonMouseMove_Bordercolor
            Frm.Controls(i).BorderColor = glButtonBordercolor
            
            Frm.Controls(i).ForeColorHover = glButtonMouseMove_Forecolor
            Frm.Controls(i).ForeColor = glButtonForecolor
            
            
'            If Frm.Controls(i).ToolTipTitle = "Spaltenanordung" Then
'                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Tabelle.jpg")
'            End If
'            If Frm.Controls(i).ToolTipTitle = "Kalender" Then
'                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Kalender.jpg")
'            End If
'
'            If Frm.Controls(i).ToolTipTitle = "Tastatur" Then
'                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Tastatur.jpg")
'            End If
            
            If Frm.Controls(i).ToolTipTitle = "Vor" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Vor.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Zurück" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "Zurück.jpg")
            End If
            
            If Frm.Controls(i).ToolTipTitle = "Wechseln" Then
                Set Frm.Controls(i).Picture = LoadPicture(cPfad & "Picture\System\" & "switch.jpg")
            End If
            
        End If
    
    Next i
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "Farbform_nurButtons"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub LoescheAltDateien(sPfad As String, Diff As Byte)
    On Error GoTo LOKAL_ERROR
    
    
    Dim lAnz        As Long
    Dim lcount      As Long
    Dim lHeute      As Long
    Dim lDateiDatum As Long
    Dim cdatei      As String
    Dim cPfad       As String
    
    lHeute = Fix(Now)
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "F*.LZH"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > Diff Then
            Kill sPfad & cdatei
        End If
    Next lcount
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "LoescheAltDateien"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function FileExists(ByVal sFile As String) As Boolean
    Dim lngSize As Long
    
    On Error Resume Next
    lngSize = FileLen(sFile)
    FileExists = (err = 0)
    On Error GoTo 0
End Function
Public Function FindFile(RootPath As String, FileName As String) As Boolean
    On Error GoTo FileFind_Error
    
    Dim cPfadFileex As String
    
    cPfadFileex = RootPath
    If Right(cPfadFileex, 1) <> "\" Then
        cPfadFileex = cPfadFileex & "\"
    End If
    
    If FileExists(cPfadFileex & FileName) Then
        FindFile = True
    Else
        FindFile = False
    End If

Exit Function
FileFind_Error:
    FindFile = False
End Function
Private Function fnVerarbeiteArtikelMOD6(dbDb As Database, picprogress As PictureBox, txtStatus As TextBox, frmx As Form) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim rsQ         As Recordset
    Dim rsZ         As Recordset
    Dim rsZin       As Recordset
    Dim rsEti       As Recordset
    Dim rsArtlief   As Recordset
    Dim sSQL        As String
    Dim rsEPRO      As Recordset
    Dim cFeldZ      As String
    Dim cSQL        As String
    Dim cSQL1       As String
    Dim bInsert     As Boolean
    Dim bStruktur   As Boolean
    Dim lKeyNr      As Long
    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim bDatenVorhanden As Boolean
    Dim j               As Integer
    Dim bBistDuPreisänderung As Boolean
    Dim bBistDuFarbänderung As Boolean
    
    bBistDuPreisänderung = False
    bBistDuFarbänderung = False
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    
    fnVerarbeiteArtikelMOD6 = 1
    
    picprogress.Visible = True
    ShowProgress picprogress, 0, 0, 0
    
    txtStatus.Text = "10"
    
    frmWKL27!Label2(1).Caption = "Filialpreise werden vorbereitet..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    loeschNEW "ZBEST_FIL", dbDb
    cSQL = "Select * into ZBEST_FIL from Zbest_IN where FILIALNR = " & CLng(gcFilNr)
    dbDb.Execute cSQL, dbFailOnError
    
    txtStatus.Text = "20"
    
    frmWKL27!Label2(1).Caption = "Filialpreise werden indiziert..."
    frmWKL27!Label2(1).Refresh
    
    cSQL = "Create Index PRIMKEY on ZBEST_FIL (FILIALNR, ARTNR)"
    dbDb.Execute cSQL, dbFailOnError
    
    txtStatus.Text = "30"
    
    frmWKL27!Label2(1).Caption = "Filialpreise erfolgreich vorbereitet"
    frmWKL27!Label2(1).Refresh
            
    cSQL = "Create Index ARTNR on ZBEST_FIL (ARTNR)"
    dbDb.Execute cSQL, dbFailOnError
    
    
    txtStatus.Text = "40"
    
    frmWKL27!Label2(1).Caption = "Artikeldaten werden indiziert(Synstatus)..."
    frmWKL27!Label2(1).Refresh
    
    cSQL = "Create Index SYNSTATUS on ART_IN (SYNSTATUS)"
    dbDb.Execute cSQL, dbFailOnError
    
    txtStatus.Text = "50"
    
    frmWKL27!Label2(1).Caption = "Artikeldaten werden indiziert(Artnr)..."
    frmWKL27!Label2(1).Refresh
    
    cSQL = "Create Index Artnr on ART_IN (Artnr)"
    dbDb.Execute cSQL, dbFailOnError
    
    txtStatus.Text = "60"
    
    frmWKL27!Label2(1).Caption = "Artikeldaten werden vorbereitet..."
    frmWKL27!Label2(1).Refresh
    
    bDatenVorhanden = True
    
    cSQL = "Select * from ART_IN where SYNSTATUS = 'D' "
    Set rsQ = dbDb.OpenRecordset(cSQL)
    
    txtStatus.Text = "70"
    
    If Not rsQ.EOF Then
        rsQ.MoveFirst
        Do While Not rsQ.EOF
        
            If Not IsNull(rsQ!artnr) Then
                lKeyNr = rsQ!artnr
            Else
                lKeyNr = 0
            End If
            
            If lKeyNr >= 545454 And lKeyNr <= 545456 Then
                lKeyNr = lKeyNr
            End If
            
            cSQL1 = "Select * from Artikel where artnr = " & lKeyNr
            Set rsZ = gdBase.OpenRecordset(cSQL1)
            If Not rsZ.EOF Then
                rsZ.delete
            Else
                
            End If
            rsZ.Close: Set rsZ = Nothing
            
            cSQL1 = "Delete from Artlief where artnr = " & lKeyNr
            gdBase.Execute cSQL1, dbFailOnError
            
        rsQ.MoveNext
        Loop
    End If
    rsQ.Close: Set rsQ = Nothing
    
    frmWKL27!Label2(1).Caption = "Artikeldaten werden überprüft..."
    frmWKL27!Label2(1).Refresh
    
    frmx.Refresh
    
    txtStatus.Text = "80"
    
    cSQL = "Delete from ART_IN where SYNSTATUS = 'D' "
    dbDb.Execute cSQL, dbFailOnError
    
    frmWKL27!Label2(1).Caption = "Artikeldaten, Anzahl wird ermittelt..."
    frmWKL27!Label2(1).Refresh
    
    txtStatus.Text = "90"
    
    If Not NewTableSuchenDBKombi("ETIPROTS", gdBase) Then
        CreateTable "ETIPROTS", gdBase
    End If
    
    If NewTableSuchenDBKombi("ETIPROT", gdBase) = True Then
    
        cSQL = "Insert into ETIPROTS Select "
        cSQL = cSQL & " ARTNR "
        cSQL = cSQL & ", BEZEICH  "
        cSQL = cSQL & ", BESTAND  "
        cSQL = cSQL & ", ANZAHL  "
        cSQL = cSQL & ", VKPRNEU  "
        cSQL = cSQL & ", VKPRALT  "
        cSQL = cSQL & ", LIBESNR  "
        cSQL = cSQL & ", EAN  "
        cSQL = cSQL & ", LPZ  "
        cSQL = cSQL & ", LINR  "
        cSQL = cSQL & ", FILNR  "
        cSQL = cSQL & ", WEDate  "
        cSQL = cSQL & ", KD  "
        cSQL = cSQL & " from ETIPROT "
        gdBase.Execute cSQL, dbFailOnError
        
    End If
        
    loeschNEW "ETIPROT", gdBase
    CreateTable "ETIPROT", gdBase
    
    cSQL = "Select * from ART_IN order by ARTNR"
    Set rsQ = dbDb.OpenRecordset(cSQL)
    
    txtStatus.Text = "100"
    
    If Not rsQ.EOF Then
        rsQ.MoveLast
        lAnzSatz = rsQ.RecordCount
        
        frmWKL27!Label2(5).Visible = True
        frmWKL27!Label2(5).Caption = Trim$(Str$(lAnzSatz))
        frmWKL27!Label2(5).Refresh
        
        Set rsEPRO = gdBase.OpenRecordset("ETIPROT", dbOpenTable)
        
        frmWKL27!Label2(1).Caption = "Artikeldaten werden eingelesen..."
        frmWKL27!Label2(1).Refresh
        
        rsQ.MoveFirst
        Do While Not rsQ.EOF
            lAktSatz = lAktSatz + 1
            
            Select Case lAnzSatz
                Case Is > 5000

                    j = lAktSatz Mod 50
                    If j = 0 Then
                        txtStatus.Text = lAktSatz * 100 / lAnzSatz
                        frmx!Label2(3).Caption = Trim$(Str$(lAktSatz))
                        frmx!Label2(3).Refresh
                        frmx.Refresh
                    End If

                Case Is > 500

                    j = lAktSatz Mod 10
                    If j = 0 Then
                        txtStatus.Text = lAktSatz * 100 / lAnzSatz
                        frmx!Label2(3).Caption = Trim$(Str$(lAktSatz))
                        frmx!Label2(3).Refresh
                        frmx.Refresh
                    End If

                Case Is <= 500
                    txtStatus.Text = lAktSatz * 100 / lAnzSatz
                    frmx!Label2(3).Caption = Trim$(Str$(lAktSatz))
                    frmx!Label2(3).Refresh
                    frmx.Refresh
            End Select
        
            If Not IsNull(rsQ!artnr) Then
                lKeyNr = rsQ!artnr
            Else
                lKeyNr = 0
            End If
            
'            If lKeyNr = 340969 Then
'                 MsgBox "hier"
'            End If
           
            
            If lKeyNr >= 545454 And lKeyNr <= 545456 Then
                lKeyNr = lKeyNr
            End If
            
            cSQL1 = "Select * from Artikel where artnr = " & lKeyNr
            Set rsZ = gdBase.OpenRecordset(cSQL1)

            If Not rsZ.EOF Then
                rsZ.Edit
                rsZ!SYNStatus = "E"
                bInsert = False
            Else
                rsZ.AddNew
                rsZ!SYNStatus = "A"
                bInsert = True
                
                
                cSQL1 = "Delete from Artlief where artnr = " & lKeyNr
                gdBase.Execute cSQL1, dbFailOnError
                
                
                
            End If
            
            '****************************************
            '* Datentransfer
            '****************************************
            
            If lKeyNr <> 0 Then
            
                'rsQ!KVKPR1   das ist der Preis aus der Artikel
                'jetzt noch checken ob bestimmter Filialpreis in der Zbest_in
                
                Dim sSQL1 As String
                Dim dkvkausZB As Double
                Dim dkvkausArt_in As Double
                Dim dPreisderGeschriebenwird As Double
                
                dkvkausZB = 0
                If Not IsNull(rsQ!KVKPR1) Then
                    dkvkausArt_in = rsQ!KVKPR1
                Else
                    dkvkausArt_in = 0
                End If
                
                Dim dMBausZB As Double
                Dim dMBausArt_in As Double
                Dim dMBderGeschriebenwird As Double
                
                dMBausZB = 0
                If Not IsNull(rsQ!MINBEST) Then
                    dMBausArt_in = rsQ!MINBEST
                Else
                    dMBausArt_in = 0
                End If
                
                sSQL1 = "Select * from ZBEST_FIL where artnr =" & lKeyNr

                Set rsZin = dbDb.OpenRecordset(sSQL1)
                If Not rsZin.EOF Then
                    If Not IsNull(rsZin!KVKPR1) Then
                        dkvkausZB = rsZin!KVKPR1
                    End If
                    
                    If Not IsNull(rsZin!MINBEST) Then
                        dMBausZB = rsZin!MINBEST
                    End If
                End If
                rsZin.Close: Set rsZin = Nothing
                
                If dkvkausZB <> 0 Then
                    If dkvkausArt_in <> dkvkausZB Then
                        'dkvkausZB spezialpreis aus Zbestand wird geschrieben
                        dPreisderGeschriebenwird = dkvkausZB
                    Else
                        'dkvkausArt_in Normalpreis aus art_in wird geschrieben
                        dPreisderGeschriebenwird = dkvkausArt_in
                    
                    End If
                Else
                    'dkvkausArt_in Normalpreis aus art_in wird geschrieben
                    dPreisderGeschriebenwird = dkvkausArt_in
                
                End If
                
                If dMBausZB <> 0 Then
                    If dMBausArt_in <> dMBausZB Then
                        'dkvkausZB spezialpreis aus Zbestand wird geschrieben
                        dMBderGeschriebenwird = dMBausZB
                    Else
                        'dkvkausArt_in Normalpreis aus art_in wird geschrieben
                        dMBderGeschriebenwird = dMBausArt_in
                    
                    End If
                Else
                    'dkvkausArt_in Normalpreis aus art_in wird geschrieben
                    dMBderGeschriebenwird = dMBausArt_in
                End If
                
                
                
                If Format(rsZ!KVKPR1, "####.00") <> Format(dPreisderGeschriebenwird, "####.00") Then 'rsQ!KVKPR1 Then
                    
                    bBistDuPreisänderung = True
                    
                    Dim bPS As Boolean
                    bPS = False
                    
                    If Not IsNull(rsZ!PREISSCHU) Then
                        If rsZ!PREISSCHU = "J" Then
                            bPS = True
                        Else
                            
                            bPS = False
                        
                        End If
                    Else
                        bPS = False
                    End If
                    
                    If bPS = False Then 'dann auch macht auch ein Etikett Sinn
                        
                        cSQL = "Select * from ETIDRU where ARTNR = " & lKeyNr
                        Set rsEti = gdBase.OpenRecordset(cSQL)
                        If Not rsEti.EOF Then
                            rsEti.delete
                        End If
                        
                        If rsZ!BESTAND > 0 Then
                            rsEti.AddNew
                            rsEti!artnr = lKeyNr
                            rsEti!BEZEICH = rsQ!BEZEICH
                            rsEti!vkpr = dPreisderGeschriebenwird 'rsQ!KVKPR1
                            rsEti!BESTAND = rsZ!BESTAND
                            rsEti!ANZAHL = rsZ!BESTAND
                            rsEti!LIBESNR = rsQ!LIBESNR
                            rsEti!EAN = rsQ!EAN
                            rsEti!linr = rsQ!linr
                            rsEti!LPZ = rsQ!LPZ
                            rsEti!filnr = CInt(gcFilNr)
                            rsEti!Pcname = srechnertab
                            rsEti.Update
                        Else
                        
                            If HatArtikelVerkäufe(lKeyNr) Then 'für Rühle
                        
                                If rsQ!GEFUEHRT = "J" Then
                                    rsEti.AddNew
                                    rsEti!artnr = lKeyNr
                                    rsEti!BEZEICH = rsQ!BEZEICH
                                    rsEti!vkpr = dPreisderGeschriebenwird 'rsQ!KVKPR1
                                    rsEti!BESTAND = 0
                                    rsEti!ANZAHL = 0
                                    rsEti!LIBESNR = rsQ!LIBESNR
                                    rsEti!EAN = rsQ!EAN
                                    rsEti!linr = rsQ!linr
                                    rsEti!LPZ = rsQ!LPZ
                                    rsEti!filnr = CInt(gcFilNr)
                                    rsEti!Pcname = srechnertab
                                    rsEti.Update
                                End If
                            End If
                        End If
                        rsEti.Close: Set rsEti = Nothing
                    
                    End If
                    
                    If rsQ!GEFUEHRT = "J" Then 'neu am 04.08.2010
                    
                        rsEPRO.AddNew
                        rsEPRO!artnr = lKeyNr
                        rsEPRO!BEZEICH = rsQ!BEZEICH
                        rsEPRO!VKPRNEU = dPreisderGeschriebenwird 'rsQ!KVKPR1
                        rsEPRO!VKPRalt = rsZ!KVKPR1
                        rsEPRO!BESTAND = rsZ!BESTAND
                        rsEPRO!ANZAHL = rsZ!BESTAND
                        rsEPRO!LIBESNR = rsQ!LIBESNR
                        rsEPRO!EAN = rsQ!EAN
                        rsEPRO!linr = rsQ!linr
                        rsEPRO!LPZ = rsQ!LPZ
                        rsEPRO!filnr = CInt(gcFilNr)
                        rsEPRO!WEDATE = DateValue(Now)
                        gbPreisAender = True
                        rsEPRO.Update
                    
                    End If
                Else
                    bBistDuPreisänderung = False
                End If
            End If
            
            rsZ!artnr = rsQ!artnr
            If gbTagAkt = True Then
                rsZ!BEZEICH = UCase(rsQ!BEZEICH)
            Else
                rsZ!BEZEICH = rsQ!BEZEICH
            End If
            rsZ!AGN = rsQ!AGN
            rsZ!PGN = rsQ!PGN
            rsZ!lekpr = rsQ!lekpr
            rsZ!vkpr = rsQ!vkpr
            rsZ!MWST = rsQ!MWST
            rsZ!linr = rsQ!linr
            rsZ!LIBESNR = rsQ!LIBESNR
            rsZ!EAN = rsQ!EAN
            rsZ!EAN2 = rsQ!EAN2
            rsZ!EAN3 = rsQ!EAN3
            rsZ!ETIMERK = rsQ!ETIMERK
            rsZ!RKZ = rsQ!RKZ
            rsZ!LPZ = rsQ!LPZ
            rsZ!NOTIZEN = Left(rsQ!NOTIZEN, 25)
            rsZ!MINMEN = rsQ!MINMEN
            rsZ!INHALT = rsQ!INHALT
            rsZ!INHALTBEZ = rsQ!INHALTBEZ
            rsZ!GRUNDPREIS = rsQ!GRUNDPREIS
            
            If Bist_du_in_Prsterm(rsQ!artnr) Then
                rsZ!RABATT_OK = "N"
            Else
                rsZ!RABATT_OK = rsQ!RABATT_OK
            End If
            
            
            rsZ!GEFUEHRT = rsQ!GEFUEHRT
            
            If bBistDuPreisänderung = True Then
                If Not IsNull(rsZ!PREISSCHU) Then
                    If rsZ!PREISSCHU = "J" Then
                        updatePrsterm rsQ!artnr, dPreisderGeschriebenwird
                    Else
                        
                        rsZ!KVKPR1 = dPreisderGeschriebenwird ' rsQ!KVKPR1 'Kassenverkauf
                        rsZ!PREISSCHU = rsQ!PREISSCHU
                    
                    End If
                Else
                    rsZ!KVKPR1 = dPreisderGeschriebenwird ' rsQ!KVKPR1 'Kassenverkauf
                    rsZ!PREISSCHU = rsQ!PREISSCHU
                End If
                
            Else
                rsZ!KVKPR1 = dPreisderGeschriebenwird ' rsQ!KVKPR1 'Kassenverkauf
                rsZ!PREISSCHU = rsQ!PREISSCHU
            
            End If
            
            rsZ!MINBEST = dMBderGeschriebenwird
            rsZ!ekpr = rsQ!ekpr 'Schnittek?
            
            rsZ!AUFDAT = rsQ!AUFDAT
            rsZ!EXDAT = rsQ!EXDAT
            rsZ!SPANNE = rsQ!SPANNE
            rsZ!BONUS_OK = rsQ!BONUS_OK
            rsZ!UMS_OK = rsQ!UMS_OK
            
            bBistDuFarbänderung = False
            If rsZ!AWM <> rsQ!AWM Then
                bBistDuFarbänderung = True
            End If
            
            rsZ!AWM = rsQ!AWM
            rsZ.Update
            rsZ.Close: Set rsZ = Nothing
            
            If gbETIBEIFARB Then
                If bBistDuFarbänderung = True Then
                    SchreibeEtiDruForEinzelArtikelnurBeiFarbÄnderung lKeyNr
                End If
            End If
            
            rsQ.MoveNext
        Loop
        rsEPRO.Close: Set rsEPRO = Nothing
    Else
        bDatenVorhanden = False
    End If
    
    rsQ.Close: Set rsQ = Nothing
    
    picprogress.Visible = False
    
    fnVerarbeiteArtikelMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteArtikelMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
   
End Function
Private Sub updatePrsterm(cART As String, dPreis As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    
    cSQL = "Update Prsterm set KVKPR1ALT = '" & dPreis & "'"
    cSQL = cSQL & " where artnr = " & cART
    
    gdBase.Execute cSQL, dbFailOnError
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "updatePrsterm"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function Bist_du_in_Prsterm(cART As String) As Boolean
On Error GoTo LOKAL_ERROR

    Bist_du_in_Prsterm = False
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from PRSTERM where artnr = " & cART & " and STATUS = 1 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        Bist_du_in_Prsterm = True
    End If
    
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "Bist_du_in_Prsterm"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function


Private Function fnVerarbeiteKundenMOD6Nacht(dbDb As Database, frmx As Form, txtStatus As TextBox, picprogress As PictureBox, pbrAbschluss As ProgressBar) As Long
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim cPfad As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    frmx.Refresh
    fnVerarbeiteKundenMOD6Nacht = 1
    picprogress.Visible = True
    txtStatus.Text = "10"
    
    frmx!Label2(3).Caption = "1"
    frmx!Label2(3).Refresh
    
    frmx!Label2(5).Caption = "10"
    frmx!Label2(5).Refresh
    
    pbrAbschluss.Max = 100
    
    loeschNEW "KUNDENTEMP", gdBase
    
    If SpalteInTabellegefundenNEW("KUN_in", "DS", dbDb) = False Then
        SpalteAnfuegenNEW "KUN_in", "DS", "BIT", dbDb
        
        cSQL = "Update KUN_in Set DS = FALSE "
        dbDb.Execute cSQL, dbFailOnError
    End If
    
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into KUNDENTEMP from KUN_in IN '" & cPfad & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    frmx.Refresh
    txtStatus.Text = "30"
    pbrAbschluss.value = 30
    
    '******
    
    frmx!Label2(3).Caption = "2"
    frmx!Label2(3).Refresh
    
    loeschNEW "KUNDENSYN", gdBase
    
    cSQL = "Select * into KUNDENSYN from KUNDEN where SYNSTATUS = 'A' "
    cSQL = cSQL & " or  SYNSTATUS = 'E' "
    cSQL = cSQL & " or  SYNSTATUS = 'D' "
    gdBase.Execute cSQL, dbFailOnError
    
    frmx.Refresh
    txtStatus.Text = "43"
    pbrAbschluss.value = 43
    
    frmx!Label2(3).Caption = "3"
    frmx!Label2(3).Refresh
    
    cSQL = "Delete from KUNDEN where KUNDNR in (select KUNDNR from KUNDENSYN) "
    gdBase.Execute cSQL, dbFailOnError
    
    '******
    frmx.Refresh
    txtStatus.Text = "55"
    pbrAbschluss.value = 55
    
    frmx!Label2(3).Caption = "4"
    frmx!Label2(3).Refresh
    
    cSQL = "Delete from KUNDEN where KUNDNR in (select KUNDNR from KUNDENTEMP) "
    gdBase.Execute cSQL, dbFailOnError
    
    frmx.Refresh
    txtStatus.Text = "69"
    pbrAbschluss.value = 69
    
    frmx!Label2(3).Caption = "5"
    frmx!Label2(3).Refresh
    
    cSQL = "Insert into KUNDEN Select * from KUNDENTEMP where KUNDNR not in (select KUNDNR from KUNDENSYN) "
    gdBase.Execute cSQL, dbFailOnError
    
    frmx.Refresh
    txtStatus.Text = "79"
    pbrAbschluss.value = 79
    
    frmx!Label2(3).Caption = "6"
    frmx!Label2(3).Refresh
    
    cSQL = "Delete from KUNDEN where status = 'D'"
    gdBase.Execute cSQL, dbFailOnError
    
    frmx.Refresh
    txtStatus.Text = "88"
    pbrAbschluss.value = 88
    
    frmx!Label2(3).Caption = "7"
    frmx!Label2(3).Refresh
    
    cSQL = "Delete from KUNDEN where synstatus = 'D'"
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = "91"
    pbrAbschluss.value = 91
    
    frmx!Label2(3).Caption = "8"
    frmx!Label2(3).Refresh
    
    cSQL = "Update KUNDEN SET synSTATUS= 'N'  where synstatus <> 'N' "
    gdBase.Execute cSQL, dbFailOnError
    

    
    frmx.Refresh
    txtStatus.Text = "95"
    pbrAbschluss.value = 95
    
    frmx!Label2(3).Caption = "9"
    frmx!Label2(3).Refresh
    
    cSQL = "Update KUNDEN SET STATUS= 'N' where STATUS <> 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    frmx.Refresh
    txtStatus.Text = "99"
    pbrAbschluss.value = 99
    
    frmx!Label2(3).Caption = "10"
    frmx!Label2(3).Refresh
    
    cSQL = "Insert into KUNDEN Select * from KUNDENsyn "
    gdBase.Execute cSQL, dbFailOnError
    
    fnVerarbeiteKundenMOD6Nacht = 0
    picprogress.Visible = False
    
Exit Function
LOKAL_ERROR:
    If err.Number = 3375 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "fnVerarbeiteKundenMOD6Nacht"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Function
Private Function fnVerarbeiteGutscheineMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim cSQL As String
    Dim rsrs1 As Recordset
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    

    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    
    fnVerarbeiteGutscheineMOD6 = 1
    
    frmWKL27.Refresh
    
    loeschNEW "Gutscht", gdBase
    cSQL = "Select * into Gutscht from Gutsch where Status <> 'N'"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "GutschN", gdBase
    cSQL = "Select * into GutschN from Gutsch where Status = 'N'"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "Gutsch", gdBase
    
    
    'KissLive Erweiterung
    If SpalteInTabellegefundenNEW("Gut_in", "DAT_AUSG_TEXT", dbDb) Then
        
        cSQL = "Update Gut_in SET DAT_AUSG_TEXT ='' where DAT_AUSG_TEXT is null "
        dbDb.Execute cSQL, dbFailOnError
        
        cSQL = "Update Gut_in SET DAT_AUSG = clng(datevalue(DAT_AUSG_TEXT)) where DAT_AUSG_TEXT  <> '' "
        dbDb.Execute cSQL, dbFailOnError
        
        cSQL = "Update Gut_in SET DAT_EINL_TEXT ='' where DAT_EINL_TEXT is null "
        dbDb.Execute cSQL, dbFailOnError
        
        cSQL = "Update Gut_in SET DAT_EINL = clng(datevalue(DAT_EINL_TEXT)) where DAT_EINL_TEXT  <> '' "
        dbDb.Execute cSQL, dbFailOnError
        
        cSQL = " Alter table Gut_in drop DAT_AUSG_TEXT "
        dbDb.Execute cSQL, dbFailOnError
        
        cSQL = " Alter table Gut_in drop DAT_EINL_TEXT "
        dbDb.Execute cSQL, dbFailOnError
    End If
    
    
    
    
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into Gutsch from Gut_in IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    '*****
    
    loeschNEW "GutschDEL", gdBase
    
    cSQL = "Select * into GutschDEL from Gutsch where STATUS= 'L' "
    gdBase.Execute cSQL, dbFailOnError
    
    
'    cSQL = "Delete from Gutsch where STATUS= 'L' "
'    gdBase.Execute cSQL, dbFailOnError
'
'    cSQL = "Delete from Gutsch where synSTATUS= 'D' "
'    gdBase.Execute cSQL, dbFailOnError
    '******
    
    cSQL = "Update GUTSCH SET STATUS= 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from GutschN where gutschnr in (select gutschnr from GUTSCH) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GUTSCH inner join GUTSCHT on GUTSCH.GUTSCHNR = GUTSCHT.GUTSCHNR "
    cSQL = cSQL & " Set gutsch.Status = gutschT.Status"
    cSQL = cSQL & " where gutschT.Status <> 'A'"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    cSQL = " Insert into GUTSCH Select * from Gutscht "
    cSQL = cSQL & " where gutschT.Status = 'A'"  'Fehler <> A
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Insert into GUTSCH Select * from GutschN "
    cSQL = cSQL & " where gutschN.Status = 'N'"
    gdBase.Execute cSQL, dbFailOnError
    
    

    
    
    
    
    

    cSQL = "Create Index DAT_EINL on GUTSCH (DAT_EINL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index GUTSCHNR on GUTSCH (GUTSCHNR)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index LastDate on GUTSCH (LastDate)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index LASTTIME on GUTSCH (Lasttime)"
    gdBase.Execute cSQL, dbFailOnError
    
    
    'Letzendlich Löschen
    cSQL = "Update GUTSCH inner join GutschDEL on GUTSCH.GUTSCHNR = GutschDEL.GUTSCHNR "
    cSQL = cSQL & " Set gutsch.Status = 'L' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from Gutsch where STATUS= 'L' "
    gdBase.Execute cSQL, dbFailOnError

            
    fnVerarbeiteGutscheineMOD6 = 0
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteGutscheineMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   Resume Next
    
End Function
Private Function fnVerarbeiteLieferantenMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim rsQ As Recordset
    Dim rsZ As Recordset
    
    Dim tdQ As TableDef
    Dim tdZ As TableDef

    Dim lAnzFelderQ As Long
    Dim lAnzFelderZ As Long
    Dim lCountQ As Long
    Dim lCountZ As Long
    Dim lcount As Long
    Dim cFeldQ As String
    Dim cFeldZ As String
    
    Dim cSQL As String
    Dim bInsert As Boolean
    Dim bStruktur As Boolean
    
    Dim lKeyNr As Long
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    
    Dim bNewFeld As Boolean
    bNewFeld = False
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    fnVerarbeiteLieferantenMOD6 = 1
    
    If SpalteInTabellegefundenNEW("LISRT", "GLN", gdBase) Then
        bNewFeld = True
    End If
    
    cSQL = "Drop Index LINR on LISRT"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index LINR on LISRT (LINR)"
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsZ = gdBase.OpenRecordset("LISRT", dbOpenTable)
    rsZ.index = "LINR"
    
    
    If SpalteInTabellegefundenNEW("LISRT_IN", "GLN", dbDb) Then
        bNewFeld = True
    Else
        bNewFeld = False
    End If
    
    cSQL = "Select * from LISRT_IN order by LINR"
    Set rsQ = dbDb.OpenRecordset(cSQL)
    If Not rsQ.EOF Then
        rsQ.MoveLast
        lAnzSatz = rsQ.RecordCount
        frmWKL27!Label2(5).Caption = Trim$(Str$(lAnzSatz))
        frmWKL27!Label2(5).Refresh
        
        rsQ.MoveFirst
        Do While Not rsQ.EOF
            lAktSatz = lAktSatz + 1
            frmWKL27!Label2(3).Caption = Trim$(Str$(lAktSatz))
            frmWKL27!Label2(3).Refresh
            
            If Not IsNull(rsQ!linr) Then
                lKeyNr = rsQ!linr
            Else
                lKeyNr = 0
            End If
            
            rsZ.Seek "=", lKeyNr
            If Not rsZ.NoMatch Then
                '****************************************
                '* Datensatz existiert -> update!
                '****************************************
                rsZ.Edit
                bInsert = False
            Else
                '****************************************
                '* Datensatz existiert nicht -> insert!
                '****************************************
                rsZ.AddNew
                bInsert = True
            End If
            
            '****************************************
            '* Datentransfer
            '****************************************

            rsZ!linr = rsQ!linr
            rsZ!Kuerzel = rsQ!Kuerzel
            rsZ!LIEFBEZ = rsQ!LIEFBEZ
            rsZ!STADT = rsQ!STADT
            rsZ!Plz = Trim(Left(rsQ!Plz, 5))
            rsZ!strasse = rsQ!strasse
            rsZ!Tel = rsQ!Tel
            rsZ!Fax = rsQ!Fax
            rsZ!Zusatz = rsQ!Zusatz
'            rsZ!Kundnr = rsQ!Kundnr 'nicht mehr abgleichen Rühle bestellt dezentral
            rsZ!NOTIZ = rsQ!NOTIZ
            rsZ!AWERT = rsQ!AWERT
            rsZ!LASTDATE = rsQ!LASTDATE
            rsZ!LASTTIME = rsQ!LASTTIME
'            rsZ!Pass = rsQ!Pass
'            rsZ!adress = rsQ!adress
'            rsZ!KennNr = rsQ!KennNr
'            rsZ!Format = rsQ!Format
'            rsZ!bUser = rsQ!bUser
            rsZ!SYNStatus = rsQ!SYNStatus
            rsZ!Email = rsQ!Email
            rsZ!KTEXT = rsQ!KTEXT
            
            If bNewFeld Then
'                rsZ!GLN = rsQ!GLN
                rsZ!USTID = rsQ!USTID
            End If
            
            rsZ.Update

            rsQ.MoveNext
        Loop
    End If
    
    rsQ.Close: Set rsQ = Nothing
    rsZ.Close: Set rsZ = Nothing
    
    fnVerarbeiteLieferantenMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    If err.Number = 3372 Or err.Number = 53 Or err.Number = 3256 Or err.Number = 3376 Or err.Number = 3043 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "fnVerarbeiteLieferantenMOD6"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1

    End If
End Function
Private Function fnVerarbeiteBedienerMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteBedienerMOD6 = 1
    
    loeschNEW "BEDNAME", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into BEDNAME from BED_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteBedienerMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteBedienerMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteAUSZAHLUNGSGRUNDMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteAUSZAHLUNGSGRUNDMOD6 = 1
    
    loeschNEW "AUSZAHLUNGSGRUND", gdBase
    CreateTableT2 "AUSZAHLUNGSGRUND", gdBase
    
    cPfad = cPfad & "zf.mdb"
    
    cSQL = "Insert into AUSZAHLUNGSGRUND Select * from AUSZAHLUNGSGRUND_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteAUSZAHLUNGSGRUNDMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteAUSZAHLUNGSGRUNDMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteMarkeMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteMarkeMOD6 = 1
    
    loeschNEW "MARKE", gdBase
    CreateTableT2 "MARKE", gdBase
    
    cPfad = cPfad & "zf.mdb"
    
    cSQL = "Insert into MARKE Select * from MARKE_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteMarkeMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteMarkeMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteBonTextMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteBonTextMOD6 = 1
    
    cPfad = cPfad & "zf.mdb"
    
    cSQL = "Select * from Bon_IN IN '" & cPfad & "'  "
    cSQL = cSQL & " where FILIALE = " & CInt(gcFilNr)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Delete from BONTEXT "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Bontext Select ZEILENNR , ZEILENTEXT  from Bon_IN IN '" & cPfad & "'  "
    cSQL = cSQL & " where FILIALE = " & CInt(gcFilNr)
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteBonTextMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteBonTextMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteWARGRUMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lartnr As Long
    Dim cSpanne As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteWARGRUMOD6 = 1
    
    cPfad = cPfad & "zf.mdb"
    
    cSQL = "Select * from WAR_IN IN '" & cPfad & "'  "
    cSQL = cSQL & " where FILIALE = " & CInt(gcFilNr)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Delete from Warengru "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Warengru "
    cSQL = cSQL & " Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", WGNR "
    cSQL = cSQL & ", BEZEICH  "
    cSQL = cSQL & ", FAKTOR "
    cSQL = cSQL & ", SGROESSE  "
    cSQL = cSQL & ", BNAME  "
    cSQL = cSQL & " from WAR_IN IN '" & cPfad & "'  "
    cSQL = cSQL & " where FILIALE = " & CInt(gcFilNr)
    gdBase.Execute cSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("WAR_IN", "SPANNE", dbDb) Then
        'Spannen verarbeiten -> in die Artikel
        
        
        cSQL = "Select Artnr,Spanne from WAR_IN IN '" & cPfad & "'  "
        cSQL = cSQL & " where FILIALE = " & CInt(gcFilNr)
        Set rsrs = dbDb.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
        
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                If Not IsNull(rsrs!artnr) Then
                    lartnr = rsrs!artnr
                Else
                    lartnr = 0
                End If
                
                
                If Not IsNull(rsrs!SPANNE) Then
                    cSpanne = rsrs!SPANNE
                Else
                    cSpanne = 0
                End If
                
                cSpanne = SwapStr(cSpanne, ",", ".")
                
                If lartnr > 0 Then
                    cSQL = "Update Artikel set Spanne = " & cSpanne & " where artnr = " & lartnr
                    gdBase.Execute cSQL, dbFailOnError
                End If
                
            rsrs.MoveNext
            Loop
            
        End If
        
        rsrs.Close: Set rsrs = Nothing
    

    End If

    fnVerarbeiteWARGRUMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteWARGRUMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteART_EANMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteART_EANMOD6 = 1
    
    cPfad = cPfad & "zf.mdb"
    
    loeschNEW "ARTEAN_K_Temp", gdBase
    
    cSQL = " Select  "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & " into ARTEAN_K_Temp from ARTEAN_IN IN '" & cPfad & "'"
    cSQL = cSQL & " where ean <> '' "
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(1 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    
    
    cSQL = "Create Index ean on ARTEAN_K_Temp (ean)"
    gdBase.Execute cSQL, dbFailOnError
    
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(2 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    
    
    loeschNEW "ARTEAN_K_DA", gdBase
    
    cSQL = "Create Table ARTEAN_K_DA"
    cSQL = cSQL & " ( "
    cSQL = cSQL & " ARTNR int "
    cSQL = cSQL & ", ean varchar(13) "
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(3 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    
    'alles rein was ean-mäßig schon bekannt ist
    
    
    
    
    
    
    cSQL = "Insert into ARTEAN_K_DA "
    cSQL = cSQL & " Select "
    cSQL = cSQL & " ARTEAN_K_Temp.ARTNR "
    cSQL = cSQL & ", ARTEAN_K_Temp.EAN "
    cSQL = cSQL & " from ARTEAN_K_Temp inner join Artikel  "
    cSQL = cSQL & " on ARTEAN_K_Temp.ean = Artikel.ean  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(4 von 10)..."
    frmWKL27!Label2(1).Refresh
    

    
    cSQL = "Insert into ARTEAN_K_DA "
    cSQL = cSQL & " Select "
    cSQL = cSQL & " ARTEAN_K_Temp.ARTNR "
    cSQL = cSQL & ", ARTEAN_K_Temp.EAN "
    cSQL = cSQL & " from ARTEAN_K_Temp inner join Artikel  "
    cSQL = cSQL & " on ARTEAN_K_Temp.ean = Artikel.ean2  "
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(5 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    cSQL = "Insert into ARTEAN_K_DA "
    cSQL = cSQL & " Select "
    cSQL = cSQL & " ARTEAN_K_Temp.ARTNR "
    cSQL = cSQL & ", ARTEAN_K_Temp.EAN "
    cSQL = cSQL & " from ARTEAN_K_Temp inner join Artikel  "
    cSQL = cSQL & " on ARTEAN_K_Temp.ean = Artikel.ean3  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(6 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    
    cSQL = "Delete * from ARTEAN_K_DA where ean is null "
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Delete * from ARTEAN_K_DA where val(ean) = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    cSQL = "Delete * from ARTEAN_K where ean = '' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete * from ARTEAN_K where ean is null "
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(7 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    cSQL = "Insert into ARTEAN_K_DA "
    cSQL = cSQL & " Select "
    cSQL = cSQL & " ARTEAN_K_Temp.ARTNR "
    cSQL = cSQL & ", ARTEAN_K_Temp.EAN "
    cSQL = cSQL & " from ARTEAN_K_Temp inner join ARTEAN_K  "
    cSQL = cSQL & " on ARTEAN_K_Temp.ean = ARTEAN_K.ean  "
    cSQL = cSQL & " where Artean_k.ean <> '' "
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(8 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    
    cSQL = "Create Index ean on ARTEAN_K_DA (ean)"
    gdBase.Execute cSQL, dbFailOnError
    
    
    'alles löschen was bekannt ist
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(9 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    cSQL = "Delete * from ARTEAN_K_Temp where ean in (Select ean from ARTEAN_K_DA)"
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL27!Label2(1).Caption = "EAN werden übernommen(10 von 10)..."
    frmWKL27!Label2(1).Refresh
    
    'alles übernehmen was über bleibt und demzufolge unbekannt ist
    
    cSQL = "Insert into ARTEAN_K "
    cSQL = cSQL & " Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & " from ARTEAN_K_Temp "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "ARTEAN_K_Temp", gdBase
    loeschNEW "ARTEAN_K_DA", gdBase

    fnVerarbeiteART_EANMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteART_EANMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Function fnVerarbeiteRezeptArtikelMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteRezeptArtikelMOD6 = 1
    
    loeschNEW "GESCHWART", gdBase
    CreateTableT2 "GESCHWART", gdBase
    
    
    cPfad = cPfad & "zf.mdb"
    cSQL = "Insert into GESCHWART Select * from REZEPTARTIKEL_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    loeschNEW "GESCHWART_TEMP", gdBase
    cSQL = "Select * into GESCHWART_TEMP from GESCHWART  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Update GESCHWART inner join GESCHWART_TEMP  "
    cSQL = cSQL & " on GESCHWART.MUTTERARTNR = GESCHWART_TEMP.MUTTERARTNR  "
    cSQL = cSQL & " and GESCHWART.ARTNR = GESCHWART_TEMP.ARTNR  "
    cSQL = cSQL & " set  GESCHWART.KVK_OK = TRUE "
    cSQL = cSQL & " where GESCHWART_TEMP.KVK_OK = FALSE"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GESCHWART inner join GESCHWART_TEMP  "
    cSQL = cSQL & " on GESCHWART.MUTTERARTNR = GESCHWART_TEMP.MUTTERARTNR  "
    cSQL = cSQL & " and GESCHWART.ARTNR = GESCHWART_TEMP.ARTNR  "
    cSQL = cSQL & " set  GESCHWART.KVK_OK = FALSE "
    cSQL = cSQL & " where GESCHWART_TEMP.KVK_OK = TRUE"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    cSQL = "Update GESCHWART inner join GESCHWART_TEMP  "
    cSQL = cSQL & " on GESCHWART.MUTTERARTNR = GESCHWART_TEMP.MUTTERARTNR  "
    cSQL = cSQL & " and GESCHWART.ARTNR = GESCHWART_TEMP.ARTNR  "
    cSQL = cSQL & " set  GESCHWART.BESTAND_OK = TRUE "
    cSQL = cSQL & " where GESCHWART_TEMP.BESTAND_OK = FALSE"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GESCHWART inner join GESCHWART_TEMP  "
    cSQL = cSQL & " on GESCHWART.MUTTERARTNR = GESCHWART_TEMP.MUTTERARTNR  "
    cSQL = cSQL & " and GESCHWART.ARTNR = GESCHWART_TEMP.ARTNR  "
    cSQL = cSQL & " set  GESCHWART.BESTAND_OK = FALSE "
    cSQL = cSQL & " where GESCHWART_TEMP.BESTAND_OK = TRUE"
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update GESCHWART inner join GESCHWART_TEMP  "
    cSQL = cSQL & " on GESCHWART.MUTTERARTNR = GESCHWART_TEMP.MUTTERARTNR  "
    cSQL = cSQL & " and GESCHWART.ARTNR = GESCHWART_TEMP.ARTNR  "
    cSQL = cSQL & " set  GESCHWART.IMETI = TRUE "
    cSQL = cSQL & " where GESCHWART_TEMP.IMETI = FALSE"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GESCHWART inner join GESCHWART_TEMP  "
    cSQL = cSQL & " on GESCHWART.MUTTERARTNR = GESCHWART_TEMP.MUTTERARTNR  "
    cSQL = cSQL & " and GESCHWART.ARTNR = GESCHWART_TEMP.ARTNR  "
    cSQL = cSQL & " set  GESCHWART.IMETI = FALSE "
    cSQL = cSQL & " where GESCHWART_TEMP.IMETI = TRUE"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    
    
    
    
    fnVerarbeiteRezeptArtikelMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteRezeptArtikelMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function fnVerarbeiteKUL_INMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteKUL_INMOD6 = 1
    
    loeschNEW "KUL_IN", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into KUL_IN from KUL_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from Kunden where kundnr in (Select kundnr from KUL_IN) "
    gdBase.Execute cSQL, dbFailOnError
    

    fnVerarbeiteKUL_INMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteKUL_INMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteBONUSL_INMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteBONUSL_INMOD6 = 1
    
    loeschNEW "BONUSL_IN", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into BONUSL_IN from BONUSL_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from KUNDENBONUS where LFNR in (Select LFNR from BONUSL_IN) "
    gdBase.Execute cSQL, dbFailOnError
    
    fnVerarbeiteBONUSL_INMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteBONUSL_INMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeitePICKLISTE_INMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeitePICKLISTE_INMOD6 = 1
    
    If NewTableSuchenDBKombi("PICKLISTE_IN", gdBase) Then
        If Datendrin("PICKLISTE_IN", gdBase) Then
        
            'dann sicher mal in eine temporäre Tabelle
            
            loeschNEW "PICKLISTE_TT", gdBase
            
            cSQL = "Select * into PICKLISTE_TT from PICKLISTE_IN "
            gdBase.Execute cSQL, dbFailOnError
            
            
            'wie immer 3x
            loeschNEW "PICKLISTE_IN", gdBase
            
            cPfad = cPfad & "zf.mdb"
            cSQL = "Select * into PICKLISTE_IN from PICKLISTE_IN IN '" & cPfad & "' where Filiale_von = " & gcFilNr
            gdBase.Execute cSQL, dbFailOnError
            'Ende, wie immer 3x
            
            cSQL = "Insert into PICKLISTE_IN Select * from PICKLISTE_TT where artnr not in (Select Artnr from PICKLISTE_IN) "
            gdBase.Execute cSQL, dbFailOnError
        
            
        Else
            loeschNEW "PICKLISTE_IN", gdBase
            
            cPfad = cPfad & "zf.mdb"
            cSQL = "Select * into PICKLISTE_IN from PICKLISTE_IN IN '" & cPfad & "' where Filiale_von = " & gcFilNr
            gdBase.Execute cSQL, dbFailOnError
        End If
    Else
        loeschNEW "PICKLISTE_IN", gdBase
    
        cPfad = cPfad & "zf.mdb"
        cSQL = "Select * into PICKLISTE_IN from PICKLISTE_IN IN '" & cPfad & "' where Filiale_von = " & gcFilNr
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    fnVerarbeitePICKLISTE_INMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeitePICKLISTE_INMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteKREDITL_INMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    Dim rsKredDel As Recordset
    Dim rsKre As Recordset
    
    Dim lartnr As Long
    Dim lMenge As Long
    Dim lKUNDNR As Long
    Dim lADate As Long
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteKREDITL_INMOD6 = 1
    
    loeschNEW "KREDITL_IN", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into KREDITL_IN from KREDITL_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from KREDITL_IN where filiale = " & CInt(gcFilNr)
    Set rsKredDel = gdBase.OpenRecordset(cSQL)
    
    If Not rsKredDel.EOF Then
        rsKredDel.MoveFirst
        Do While Not rsKredDel.EOF
        
            lartnr = 0
            lMenge = 0
            lKUNDNR = 0
            lADate = 0
            
            If Not IsNull(rsKredDel!Menge) Then
                lMenge = rsKredDel!Menge
            End If
            
            If Not IsNull(rsKredDel!artnr) Then
                lartnr = rsKredDel!artnr
            End If
            
            If Not IsNull(rsKredDel!ADATE) Then
                lADate = rsKredDel!ADATE
            End If
            
            If Not IsNull(rsKredDel!Kundnr) Then
                lKUNDNR = rsKredDel!Kundnr
            End If
            
            
            
            cSQL = "Select * from KREDIT where artnr = " & lartnr
            cSQL = cSQL & " and Menge = " & lMenge
            cSQL = cSQL & " and kundnr = " & lKUNDNR
            cSQL = cSQL & " and adate = " & lADate
            Set rsKre = gdBase.OpenRecordset(cSQL)
            If Not rsKre.EOF Then
                rsKre.delete
            Else
                
            End If
            rsKre.Close: Set rsKre = Nothing
            
            
            
        rsKredDel.MoveNext
        Loop
    End If
    rsKredDel.Close: Set rsKredDel = Nothing
    
    
    
    fnVerarbeiteKREDITL_INMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteKREDITL_INMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteFARBMERKMALMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "zf.mdb"
    
    fnVerarbeiteFARBMERKMALMOD6 = 1
    
    loeschNEW "FARBMERK", gdBase
    cSQL = "Select * into FARBMERK from FARB_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "FARBKU", gdBase
    cSQL = "Select * into FARBKU from FARBK_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteFARBMERKMALMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteFARBMERKMALMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteZUORDEANMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "zf.mdb"
    
    fnVerarbeiteZUORDEANMOD6 = 1
    
    loeschNEW "ZUORDEAN", gdBase
    cSQL = "Select * into ZUORDEAN from ZUORDEAN_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteZUORDEANMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteZUORDEANMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteSTORNOFMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "zf.mdb"
    
    fnVerarbeiteSTORNOFMOD6 = 1
    
    loeschNEW "STORNOF", gdBase
    cSQL = "Select * into STORNOF from STORNOF_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteSTORNOFMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteSTORNOFMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeitePSMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "zf.mdb"
    
    fnVerarbeitePSMOD6 = 1
    
    cSQL = "Select * from PS_IN IN '" & cPfad & "'  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cSQL = "Update Artikel set Preisschu = 'N' where artnr = " & rsrs!artnr
                gdBase.Execute cSQL, dbFailOnError
            End If
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    

    fnVerarbeitePSMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeitePSMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteZBRESTMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteZBRESTMOD6 = 1
    
    loeschNEW "ZBREST", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into ZBREST from ZBREST_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteZBRESTMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteZBRESTMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteUMZUGARTIKELMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim rsBESTAND As DAO.Recordset
    Dim lBestand As Long
    
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteUMZUGARTIKELMOD6 = 1
    
    If NewTableSuchenDBKombi("UMZUGARTIKEL", gdBase) = False Then
        CreateTableT2 "UMZUGARTIKEL", gdBase
    End If
    
    cPfad = cPfad & "zf.mdb"
    cSQL = "Insert into UMZUGARTIKEL Select * from UMZUGARTIKEL_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    Dim iDEL_Artnr As Long
    Dim iNEU_Artnr As Long
    
    cSQL = "Select * from UMZUGARTIKEL_IN IN '" & cPfad & "'  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!DEL_ARTNR) Then
                iDEL_Artnr = rsrs!DEL_ARTNR
            End If
            
            If Not IsNull(rsrs!NEU_ARTNR) Then
                iNEU_Artnr = rsrs!NEU_ARTNR
            End If
            
            
            If iNEU_Artnr <> 0 And iDEL_Artnr <> 0 Then
            
                'Verkaufsdaten
                cSQL = "Update Kassjour set Artnr  = " & iNEU_Artnr & " where artnr = " & iDEL_Artnr
                gdBase.Execute cSQL, dbFailOnError
                
                'Bestandsdaten
                
                lBestand = 0
                cSQL = "Select Bestand from Artikel where artnr = " & iDEL_Artnr
                Set rsBESTAND = gdBase.OpenRecordset(cSQL)
                If Not rsBESTAND.EOF Then
                    If Not IsNull(rsBESTAND!BESTAND) Then
                        lBestand = rsBESTAND!BESTAND
                    End If
                End If
                rsBESTAND.Close: Set rsBESTAND = Nothing
                
                cSQL = "Update Artikel set Bestand = 0 where bestand is null and artnr = " & iNEU_Artnr
                gdBase.Execute cSQL, dbFailOnError
                cSQL = "Update Artikel set Bestand = Bestand +  " & lBestand & " where artnr = " & iNEU_Artnr
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Artikel set Bestand = 0 where artnr = " & iDEL_Artnr
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update ZBESTAND set MINBEST = 0 where artnr = " & iDEL_Artnr
                gdBase.Execute cSQL, dbFailOnError
                
            End If
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    

    fnVerarbeiteUMZUGARTIKELMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteUMZUGARTIKELMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function

Private Function fnVerarbeiteZUNTERMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteZUNTERMOD6 = 1
    
    loeschNEW "ZUNTER", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into ZUNTER from ZUNTER_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteZUNTERMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteZUNTERMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteZBLOCKMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteZBLOCKMOD6 = 1
    
    loeschNEW "ZBLOCK", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into ZBLOCK from ZBLOCK_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteZBLOCKMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteZBLOCKMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteZSPERRMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteZSPERRMOD6 = 1
    
    loeschNEW "ZSPERR", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into ZSPERR from ZSPERR_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteZSPERRMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteZSPERRMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteZBONUSRUECKMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteZBONUSRUECKMOD6 = 1
    
    
    cPfad = cPfad & "zf.mdb"
    
    'Arbeitstabelle
    
    loeschNEW "ARB_KUNDENBONUS", gdBase
    cSQL = "Select * into ARB_KUNDENBONUS from KUNDENBONUS_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    'Update mit denen die schon enthalten sind
    
    cSQL = "Update KUNDENBONUS inner join ARB_KUNDENBONUS  "
    cSQL = cSQL & " on KUNDENBONUS.lfnr = ARB_KUNDENBONUS.lfnr  "
    cSQL = cSQL & " set KUNDENBONUS.EINGELOEST_DATUM = ARB_KUNDENBONUS.EINGELOEST_DATUM "
    cSQL = cSQL & " , KUNDENBONUS.SENDOK = TRUE "
    gdBase.Execute cSQL, dbFailOnError
    
    
    'Insert mit denen die noch nicht enthalten sind
    
    cSQL = "Insert into KUNDENBONUS Select "
    cSQL = cSQL & " lfnr"
    cSQL = cSQL & ", KUNDNR "
    cSQL = cSQL & ", DATUM  "
    cSQL = cSQL & ", AUSZAHLBONUS  "
    cSQL = cSQL & ", BISHERIGER_BONUS  "
    cSQL = cSQL & ", EINGELOEST_DATUM  "
    cSQL = cSQL & ", True as SENDOK "
    cSQL = cSQL & " from ARB_KUNDENBONUS"
    cSQL = cSQL & " where not lfnr in (Select lfnr from Kundenbonus)"
    gdBase.Execute cSQL, dbFailOnError
    


    loeschNEW "ARB_KUNDENBONUS", gdBase

    fnVerarbeiteZBONUSRUECKMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteZBONUSRUECKMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteZBONUS_SYS(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteZBONUS_SYS = 1
    
    
    cPfad = cPfad & "zf.mdb"
    
    'Arbeitstabelle
    
    loeschNEW "ARB_BONUS", gdBase
    cSQL = "Select * into ARB_BONUS from BONUS_SYS_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    'Update mit denen die schon enthalten sind
    
    cSQL = "Update BONUS_SYS inner join ARB_BONUS  "
    cSQL = cSQL & " on BONUS_SYS.bonus_nr = ARB_BONUS.bonus_nr and BONUS_SYS.bonus_AusgabeFil = ARB_BONUS.bonus_AusgabeFil "
    cSQL = cSQL & " set BONUS_SYS.BONUS_EINLDAT = BONUS_SYS.BONUS_EINLDAT "
    cSQL = cSQL & " , BONUS_SYS.BONUS_EINLZEIT = BONUS_SYS.BONUS_EINLZEIT "
    cSQL = cSQL & " , BONUS_SYS.SENDOK = TRUE "
    gdBase.Execute cSQL, dbFailOnError
    
    
    'Insert mit denen die noch nicht enthalten sind
    
    cSQL = "Insert into BONUS_SYS Select "
    cSQL = cSQL & " BONUS_NR "
    cSQL = cSQL & ", BONUS_BETRAG "
    cSQL = cSQL & ", BONUS_AUSGABEDAT "
    cSQL = cSQL & ", BONUS_AUSGABEZEIT "
    cSQL = cSQL & ", BONUS_EINLDAT "
    cSQL = cSQL & ", BONUS_EINLZEIT "
    cSQL = cSQL & ", BONUS_AUSGABEFIL "
    cSQL = cSQL & ", True as SENDOK "
    cSQL = cSQL & " from ARB_BONUS"
    cSQL = cSQL & " where not BONUS_NR in (Select BONUS_SYS.BONUS_NR from BONUS_SYS inner join ARB_BONUS on "
    cSQL = cSQL & " BONUS_SYS.bonus_AusgabeFil = ARB_BONUS.bonus_AusgabeFil) "
    gdBase.Execute cSQL, dbFailOnError
    


    loeschNEW "ARB_KUNDENBONUS", gdBase

    fnVerarbeiteZBONUS_SYS = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteZBONUS_SYS"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteKUKASSMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteKUKASSMOD6 = 1
    
    cPfad = cPfad & "zf.mdb"
    
    cSQL = "Insert into KundKass Select *  from KUKASS_IN IN '" & cPfad & "'  "
    cSQL = cSQL & " where Filiale <> " & CInt(gcFilNr)
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteKUKASSMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteKUKASSMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeitePRSTERMMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeitePRSTERMMOD6 = 1
    
    cPfad = cPfad & "zf.mdb"
    
    cSQL = "Insert into PRSTERM Select   "
    cSQL = cSQL & " artnr "
    cSQL = cSQL & ", KVKPR1ALT "
    cSQL = cSQL & ", KVKPR1NEU "
    cSQL = cSQL & ", DAT_VON "
    cSQL = cSQL & ", DAT_BIS "
    cSQL = cSQL & ", Filiale "
    cSQL = cSQL & ", STATUS "
    cSQL = cSQL & ", RABATT_OK "
    cSQL = cSQL & ", BONUS_OK "
    cSQL = cSQL & ", PREISSCHU "
    cSQL = cSQL & ", PREISNR "
    cSQL = cSQL & "  from PRSTERM_IN IN '" & cPfad & "' where Filiale = " & CInt(gcFilNr)
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeitePRSTERMMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeitePRSTERMMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeitePREISTERMMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeitePREISTERMMOD6 = 1
    
    cPfad = cPfad & "zf.mdb"
    
    cSQL = "Insert into PREISTERM Select  "
    cSQL = cSQL & " PREISNAME "
    cSQL = cSQL & ", PREISNR "
    cSQL = cSQL & ", PREISBESCH "
    cSQL = cSQL & ", VON "
    cSQL = cSQL & ", BIS "
    cSQL = cSQL & " from PREISTERM_IN IN '" & cPfad & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from PREISTERM where Preisnr not in (select preisnr from PRSTERM) "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeitePREISTERMMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeitePREISTERMMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteFILIALENMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteFILIALENMOD6 = 1
    
    loeschNEW "FILIALEN", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into FILIALEN from FIL_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteFILIALENMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteFILIALENMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteBEDZUGRIMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteBEDZUGRIMOD6 = 1
    
    loeschNEW "BEDZUGRI", gdBase
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into BEDZUGRI from ZUG_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    gbZugriffNew = True
    cSQL = "Update DBEINSTE Set ZUGRIFFNEW = true"
    gdBase.Execute cSQL, dbFailOnError

    fnVerarbeiteBEDZUGRIMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteBEDZUGRIMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnVerarbeiteUMS_ARTFMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    '***********************************************************
    '* Rückgabewert der Funktion ist primär auf Fehler gesetzt.
    '* Erst am Ende der Funktion wird Okay gegeben
    '***********************************************************
    
    fnVerarbeiteUMS_ARTFMOD6 = 1
    
    Dim cPfad As String
    Dim cSQL As String
    Dim lMonat As Long
    Dim lJahr As Long
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "zf.mdb"
    
    loeschNEW "UARTF", gdBase
    cSQL = "Select * into UARTF from UARTF_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index Monat on UARTF (Monat)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index Jahr on UARTF (Jahr)"
    gdBase.Execute cSQL, dbFailOnError
    
    lMonat = ermMonat
    lJahr = ermJahr
    
    If Not NewTableSuchenDBKombi("UMS_ARTF", gdBase) Then
        CreateTable "UMS_ARTF", gdBase
    End If
    
    If lMonat <> 0 Then
        If lJahr <> 0 Then
    
        cSQL = "Delete from UMS_ARTF where Monat = " & lMonat
        cSQL = cSQL & " and JAHR = " & lJahr
        gdBase.Execute cSQL, dbFailOnError

        cSQL = "Insert into UMS_ARTF Select * from UARTF  "
        gdBase.Execute cSQL, dbFailOnError
        
        End If
    End If

    fnVerarbeiteUMS_ARTFMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "fnVerarbeiteUMS_ARTFMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function ermMonat() As Long
    On Error GoTo LOKAL_ERROR
    
    ermMonat = 0
    
    Dim rs As Recordset
    Dim cSQL As String
    
    cSQL = "Select Max(Monat) as maxi from UARTF"
    Set rs = gdBase.OpenRecordset(cSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            ermMonat = rs!maxi
        End If
    End If

    rs.Close: Set rs = Nothing
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "ermMonat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function ermJahr() As Long
    On Error GoTo LOKAL_ERROR
    
    ermJahr = 0
    
    Dim rs As Recordset
    Dim cSQL As String
    
    cSQL = "Select Max(Jahr) as maxi from UARTF"
    Set rs = gdBase.OpenRecordset(cSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            ermJahr = rs!maxi
        End If
    End If
    rs.Close: Set rs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "ermJahr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Function fnVerarbeiteLinienMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteLinienMOD6 = 1
    
    loeschNEW "LINBEZ", gdBase

    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into LINBEZ from LINB_in IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("LINBEZ", "SORTI", gdBase) = False Then
        SpalteAnfuegenNEW "LINBEZ", "SORTI", "INTEGER", gdBase
        
        cSQL = "Update LINBEZ Set SORTI = LPZ "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
   
    
    fnVerarbeiteLinienMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "fnVerarbeiteLinienMOD6"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Function
Private Function fnVerarbeite_Bestellungen_Export_MOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeite_Bestellungen_Export_MOD6 = 1
    
    loeschNEW "Bestellungen_EPX", gdBase

    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into Bestellungen_EPX from Bestellungen_Export_in IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    Bestellungen_Export_import
    
    fnVerarbeite_Bestellungen_Export_MOD6 = 0
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "fnVerarbeite_Bestellungen_Export_MOD6"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Function
Private Sub Bestellungen_Export_import()
On Error GoTo LOKAL_ERROR

    If NewTableSuchenDBKombi("Bestellungen_EPX", gdBase) = False Then
        Exit Sub
    End If
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    Dim lLinr As Long
    Dim lcount As Long
    Dim bgefunden As Boolean
    Dim cZiel As String
    Dim lDatum As Long
    Dim cDatum As String
    
    lDatum = Fix(Now)
    cDatum = Trim$(Str$(lDatum))
    
    
    sSQL = "Delete * from Bestellungen_EPX where Filiale <> " & gcFilNr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select distinct(linr) from Bestellungen_EPX  "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            lLinr = 0
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
                
                'Namen zuweisen
                lcount = 65
                bgefunden = True
                Do While bgefunden
                    cZiel = "Q" & lLinr & Chr$(lcount)
                    If NewTableSuchenDBKombi("Q" & lLinr & Chr$(lcount), gdBase) Then
                        bgefunden = True
                        lcount = lcount + 1
                    Else
                        bgefunden = False
                    End If
                Loop
                
                If lcount > 89 Then
                    Dim ctempa As String
                    ctempa = "Bestellvorschlag speichern nicht möglich. Die Vergabe eines Dateinamens ist gescheitert." & vbCrLf
                    ctempa = ctempa & "Löschen Sie erledigte Bestellungen im Wareneingang aus Bestellung!"
                    MsgBox ctempa, vbOKOnly + vbInformation, "Winkiss Hinweis:"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            
                loeschNEW cZiel, gdBase

                sSQL = "Select * into " & cZiel & " from Bestellungen_EPX "
                sSQL = sSQL & " where BESTVOR > 0 "
                sSQL = sSQL & " and  linr = " & lLinr
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Delete * from Bestellungen_EPX "
                sSQL = sSQL & " where linr = " & lLinr
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Alter table " & cZiel & " DROP Filiale "
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Alter table " & cZiel & " DROP LFBESTELLNR "
                gdBase.Execute sSQL, dbFailOnError
                
                'BESTREST füllen
                sSQL = "Delete from BESTREST where DATEINAME = '" & cZiel & ".DBF'"
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Insert into BESTREST "
                sSQL = sSQL & "Select LINR, "
                sSQL = sSQL & "ARTNR, LEKPR, BESTVOR, '" & cZiel & ".DBF' as DATEINAME, "
                sSQL = sSQL & cDatum & " as BEST_DATUM, " & cDatum & " as UPD_DATUM "
                sSQL = sSQL & " from " & cZiel & " where BESTVOR <> 0 "
                gdBase.Execute sSQL, dbFailOnError

            End If
            rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
            
    
    
            
    
err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "Bestellungen_Export_import"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnVerarbeiteArtGruppenMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    fnVerarbeiteArtGruppenMOD6 = 1
    
    loeschNEW "AGNDBF", gdBase
    
    cPfad = cPfad & "zf.mdb"
    cSQL = "Select * into AGNDBF from AGN_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
        
    gdBase.TableDefs.Refresh 'Dabarefresh
    
    fnVerarbeiteArtGruppenMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "fnVerarbeiteArtGruppenMOD6"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Function
Private Function fnVerarbeiteArtliefMOD6(dbDb As Database) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad       As String
    Dim cSQL        As String
    Dim rsArtlief   As Recordset
    Dim rsQ         As Recordset
    Dim lartnr      As Long
    Dim lLinr       As Long
    
    fnVerarbeiteArtliefMOD6 = 1
    
    frmWKL27!Label2(1).Caption = "Löschweitergabe wird beachtet..."
    frmWKL27!Label2(1).Refresh
    
    If SpalteInTabellegefundenNEW("ARTL_IN", "EXDAT", dbDb) = False Then
        SpalteAnfuegenNEW "ARTL_IN", "EXDAT", "DATETIME", dbDb
        SpalteAnfuegenNEW "ARTL_IN", "RKZ", "Text(1)", dbDb
    End If

    cSQL = "Select * from ARTL_IN where SYNSTATUS = 'D' "
    Set rsQ = dbDb.OpenRecordset(cSQL)
    If Not rsQ.EOF Then
        rsQ.MoveFirst
        Do While Not rsQ.EOF
            If Not IsNull(rsQ!artnr) Then
                lartnr = rsQ!artnr
                If Not IsNull(rsQ!linr) Then
                    lLinr = rsQ!linr
                    
                    cSQL = " Delete from artlief where artnr = " & lartnr & " and LINR = " & lLinr
                    gdBase.Execute cSQL, dbFailOnError
                End If
            End If
        rsQ.MoveNext
        Loop
    End If
    rsQ.Close: Set rsQ = Nothing
    
    cSQL = " Delete from ARTL_IN where SYNSTATUS = 'D' "
    dbDb.Execute cSQL, dbFailOnError
    
    
    'Neu
    
    frmWKL27!Label2(1).Caption = "Artikel-Lieferanten werden geprüft..."
    frmWKL27!Label2(1).Refresh
    
    loeschNEW "ARTL_IN" & srechnertab, dbDb
    cSQL = " Select * into ARTL_IN" & srechnertab & " from ARTL_IN "
    dbDb.Execute cSQL, dbFailOnError
    
    SpalteAnfuegenNEW "ARTL_IN" & srechnertab, "erkannt", "Text(1)", dbDb
    
    
    Dim sAccPfad As String
    sAccPfad = gcDBPfad & "\kissdata.mdb"

    loeschNEW "ARTL_IN" & srechnertab, gdBase
    TransferTab dbDb, sAccPfad, "ARTL_IN" & srechnertab
    
    
    'neu
    cSQL = "Delete * from Artlief where Artnr in (Select artnr from ARTL_IN" & srechnertab & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL27!Label2(1).Caption = "neue Artikel-Lieferanten werden angefügt..."
    frmWKL27!Label2(1).Refresh

    cSQL = "Insert into  Artlief select  "
    cSQL = cSQL & " ARTL_IN" & srechnertab & ".artnr  "
    cSQL = cSQL & " ,ARTL_IN" & srechnertab & ".linr  "
    cSQL = cSQL & " ,ARTL_IN" & srechnertab & ".LIBESNR "
    cSQL = cSQL & " ,ARTL_IN" & srechnertab & ".lekpr "
    cSQL = cSQL & " ,ARTL_IN" & srechnertab & ".MINMEN "
    cSQL = cSQL & " ,ARTL_IN" & srechnertab & ".EXDAT "
    cSQL = cSQL & " ,ARTL_IN" & srechnertab & ".RKZ "
    cSQL = cSQL & " from ARTL_IN" & srechnertab & " "
    gdBase.Execute cSQL, dbFailOnError
    

    loeschNEW "ARTL_IN" & srechnertab, gdBase
    
    'Neu Ende

    fnVerarbeiteArtliefMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "fnVerarbeiteArtliefMOD6"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Private Function fnVerarbeiteZBestandMOD6(dbDb As Database, txtStatus As TextBox, picprogress As PictureBox) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cSQL As String

    fnVerarbeiteZBestandMOD6 = 1

    picprogress.Visible = True
    ShowProgress picprogress, 0, 0, 0


    frmWKL27!Label2(1).Caption = "Bestandsdaten werden vorbereitet..."
    frmWKL27!Label2(1).Refresh

    loeschNEW "zbestand", gdBase
    CreateTable "ZBESTAND", gdBase
    
    'neu
    
    cPfad = gsKinPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "zf.mdb"
    
    cSQL = "Insert into ZBESTAND Select "
    cSQL = cSQL & " artnr "
    cSQL = cSQL & " ,FILIALNR "
    cSQL = cSQL & " ,BESTAND "
    cSQL = cSQL & " ,MINBEST "
    cSQL = cSQL & " ,KVKPR1 "
    cSQL = cSQL & " from ZBEST_IN IN '" & cPfad & "'  "
    gdBase.Execute cSQL, dbFailOnError
    
    'neu Ende

    txtStatus.Text = "30"
    frmWKL27!Label2(1).Caption = "Bestandsdaten werden indiziert(FILIALNR, ARTNR)..."
    frmWKL27!Label2(1).Refresh

    cSQL = "Create Index PRIMKEY on ZBESTAND (FILIALNR, ARTNR)"
    gdBase.Execute cSQL, dbFailOnError

    txtStatus.Text = "50"
    frmWKL27!Label2(1).Caption = "Bestandsdaten werden indiziert(ARTNR)..."
    frmWKL27!Label2(1).Refresh

    cSQL = "Create Index ARTNR on ZBESTAND (ARTNR)"
    gdBase.Execute cSQL, dbFailOnError

    txtStatus.Text = "70"
    frmWKL27!Label2(1).Caption = "Bestandsdaten werden indiziert(LastDate)..."
    frmWKL27!Label2(1).Refresh

    cSQL = "Create Index LastDate on ZBESTAND(LastDate)"
    gdBase.Execute cSQL, dbFailOnError

    If gbBestinZ Then
        cSQL = "Update Artikel set bestand = 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update Artikel inner join zbestand on Artikel.artnr = zbestand.artnr "
        cSQL = cSQL & " set artikel.bestand = zbestand.bestand where zbestand.filialnr = " & gcFilNr
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    'wenn Kisslive und Winkiss( wie bei Bingger )
    'Kisslive gibt Bestandsführung an, so dürfen die Winkiss - Bestände nicht ins Kisslive zurück-
    'geschrieben werden
    'aber die Winkiss- Bestände müssen mit dem Einlesen der Y-Dateien aktualisiert werden
    
    If gbKL_LIVEBESTAND_DIFF = True Then
        cSQL = "Update Artikel set bestand = 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update Artikel inner join zbestand on Artikel.artnr = zbestand.artnr "
        cSQL = cSQL & " set artikel.bestand = zbestand.bestand where zbestand.filialnr = " & gcFilNr
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    picprogress.Visible = False
    
    fnVerarbeiteZBestandMOD6 = 0
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul6"
        Fehler.gsFunktion = "fnVerarbeiteZBestandMOD6"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Function SucheFilialKassenDateienMOD6(picprogress As PictureBox, txtStatus As TextBox, KIcpfad As String, iLoeschen As Integer, pbrAbschluss As ProgressBar, frmx As Form) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim dbDb        As Database
    Dim cdabapfad   As String
    Dim rsQ         As Recordset
    Dim lRet        As Long
    Dim iFileNr     As Integer
    Dim lcount      As Long
    Dim lAnzTable   As Long
    Dim sName       As String
    Dim sSQL        As String
    Dim iRet        As Integer
    
    SucheFilialKassenDateienMOD6 = False
    
    pbrAbschluss.Visible = True
    pbrAbschluss.Max = 2000
    pbrAbschluss.value = 100
    
    cdabapfad = gcDBPfad
    If Right(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "kissdata.mdb"
    
    '*************************************************************
    '* Öffne Datenbank-Verzeichnis
    '*************************************************************
    
    Set dbDb = OpenDatabase(KIcpfad & "ZF.mdb", False)
    
    '*************************************************************
    '* Prüfe, ob die jeweilige Datenbank-Tabelle vorhanden ist
    '*************************************************************
    
    '*************************************************************
    '* Schritt 1: PS_IN.DBF (Preischutz von Artikeln aufheben)
    '************************************************************
   
   
    frmWKL27!Label2(1).Caption = "Artikeldaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh

    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh

    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh

    If NewTableSuchenDBKombi("PS_IN", dbDb) Then
        lRet = fnVerarbeitePSMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "PS_IN", dbDb
            End If
        End If
    End If
    
    
    
    '*************************************************************
    '* Schritt 1: ART_IN.DBF (ARTIKEL)
    '************************************************************
   
   
    frmWKL27!Label2(1).Caption = "Artikeldaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh

    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh

    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh

    If NewTableSuchenDBKombi("ART_IN", dbDb) Then
        lRet = fnVerarbeiteArtikelMOD6(dbDb, picprogress, txtStatus, frmx)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ART_IN", dbDb
            End If
        End If
    End If
    
    '*************************************************************
    '* Schritt 1b: UMZUGARTIKEL_IN
    '*************************************************************
    
    pbrAbschluss.value = 150
    frmWKL27!Label2(1).Caption = "Umzugsartikel werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("UMZUGARTIKEL_IN", dbDb) Then
        lRet = fnVerarbeiteUMZUGARTIKELMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "UMZUGARTIKEL_IN", dbDb
            End If
        End If
    Else
    
    End If

    '*************************************************************
    '* Schritt 2: KUN_IN.DBF (KUNDEN)
    '*************************************************************
    pbrAbschluss.value = 200
    frmWKL27!Label2(1).Caption = "Kundendaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("KUN_IN", dbDb) Then

        lRet = fnVerarbeiteKundenMOD6Nacht(dbDb, frmx, txtStatus, picprogress, pbrAbschluss)

        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "KUN_IN", dbDb
            End If
        End If
    Else

    End If

    '*************************************************************
    '* Schritt 2: UARTF_IN
    '*************************************************************
    
    pbrAbschluss.Max = 2000
    pbrAbschluss.value = 300
    frmWKL27!Label2(1).Caption = "Artikelumsätze werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("UARTF_IN", dbDb) Then
        lRet = fnVerarbeiteUMS_ARTFMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "UARTF_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 2: FARB_IN
    '*************************************************************
    pbrAbschluss.value = 400
    frmWKL27!Label2(1).Caption = "Farbmerkmale werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("FARB_IN", dbDb) Then
        lRet = fnVerarbeiteFARBMERKMALMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "FARB_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 2: ZUORDEAN_IN
    '*************************************************************
    pbrAbschluss.value = 450
    frmWKL27!Label2(1).Caption = "Umverpackungen werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("ZUORDEAN_IN", dbDb) Then
        lRet = fnVerarbeiteZUORDEANMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ZUORDEAN_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    '*************************************************************
    '* Schritt 2: ZUORDEAN_IN
    '*************************************************************
    pbrAbschluss.value = 475
    frmWKL27!Label2(1).Caption = "Stornoinfos werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("STORNOF_IN", dbDb) Then
        lRet = fnVerarbeiteSTORNOFMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "STORNOF_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    
    
    
    
    
    
    
    
    If gbKL_LIVEGUTSCHEIN Then
        'bei gutschein live nichts machen
    Else

        '*************************************************************
        '* Schritt 3: GUT_IN.DBF (GUTSCHEINE)
        '*************************************************************
        pbrAbschluss.value = 500
        frmWKL27!Label2(1).Caption = "Gutscheindaten werden synchronisiert..."
        frmWKL27!Label2(1).Refresh
        
        frmWKL27!Label2(3).Caption = "0"
        frmWKL27!Label2(3).Refresh
        
        frmWKL27!Label2(5).Caption = "0"
        frmWKL27!Label2(5).Refresh
        
    
        If NewTableSuchenDBKombi("GUT_IN", dbDb) Then
            lRet = fnVerarbeiteGutscheineMOD6(dbDb)
            If lRet = 0 Then
                If iLoeschen = vbYes Then
                    loeschNEW "GUT_IN", dbDb
                End If
            End If
        Else
        
        End If
    
    End If
    
    '*************************************************************
    '* Schritt 3a: KUL_IN.DBF (Kunden löschen)
    '*************************************************************
    pbrAbschluss.value = 500
    frmWKL27!Label2(1).Caption = "Kunden werden gelöscht..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    

    If NewTableSuchenDBKombi("KUL_IN", dbDb) Then
        lRet = fnVerarbeiteKUL_INMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "KUL_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 4: LISRT.DBF (LIEFERANTEN)
    '*************************************************************
    pbrAbschluss.value = 600
    frmWKL27!Label2(1).Caption = "Lieferantendaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    

    If NewTableSuchenDBKombi("LISRT_IN", dbDb) Then
        lRet = fnVerarbeiteLieferantenMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "LISRT_IN", dbDb
            End If
        End If
    End If

    '*************************************************************
    '* Schritt 5: LINB_IN.DBF (LINIEN)
    '*************************************************************
    pbrAbschluss.value = 650
    frmWKL27!Label2(1).Caption = "Liniendaten werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
 
    If NewTableSuchenDBKombi("LINB_IN", dbDb) Then
        lRet = fnVerarbeiteLinienMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "LINB_IN", dbDb
            End If
        End If
    Else
    
    End If

    
    '*************************************************************
    '* Schritt 6: ZBEST_IN.DBF (ZENTRALBESTAND)
    '*************************************************************
    pbrAbschluss.value = 700
    frmWKL27!Label2(1).Caption = "Bestandsdaten werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    

    If NewTableSuchenDBKombi("ZBEST_IN", dbDb) Then
        lRet = fnVerarbeiteZBestandMOD6(dbDb, txtStatus, picprogress)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ZBEST_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 7: BED_IN (Bediener)
    '*************************************************************
    pbrAbschluss.value = 750
    frmWKL27!Label2(1).Caption = "Bedienerdaten werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    

    If NewTableSuchenDBKombi("BED_IN", dbDb) Then
        lRet = fnVerarbeiteBedienerMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "BED_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 7: AUSZAHLUNGSGRUND_IN (AUSZAHLUNGSGRUND)
    '*************************************************************
    pbrAbschluss.value = 800
    frmWKL27!Label2(1).Caption = "Auszahlungsgründe werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    
    If NewTableSuchenDBKombi("AUSZAHLUNGSGRUND_IN", dbDb) Then
        lRet = fnVerarbeiteAUSZAHLUNGSGRUNDMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "AUSZAHLUNGSGRUND_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 7: Marke_IN (Rabattgrenze)
    '*************************************************************
    pbrAbschluss.value = 850
    frmWKL27!Label2(1).Caption = "Rabattgrenzen werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    
    If NewTableSuchenDBKombi("MARKE_IN", dbDb) Then
        lRet = fnVerarbeiteMarkeMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "MARKE_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    pbrAbschluss.value = 900
    
    frmWKL27!Label2(1).Caption = "Warengruppen werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    
    If NewTableSuchenDBKombi("WAR_IN", dbDb) Then
        lRet = fnVerarbeiteWARGRUMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "WAR_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    
    pbrAbschluss.value = 950
    frmWKL27!Label2(1).Caption = "EAN werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    
    If NewTableSuchenDBKombi("ARTEAN_IN", dbDb) Then
        lRet = fnVerarbeiteART_EANMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ARTEAN_IN", dbDb
            End If
        End If
    Else
    
    End If

    
    pbrAbschluss.value = 1000
    
    frmWKL27!Label2(1).Caption = "Rezeptartikel werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    
    If NewTableSuchenDBKombi("REZEPTARTIKEL_IN", dbDb) Then
        lRet = fnVerarbeiteRezeptArtikelMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "REZEPTARTIKEL_IN", dbDb
            End If
        End If
    Else
    
    End If

    
    '*************************************************************
    '* Schritt 8: AGNDBF.DBF (ARTIKELGRUPPEN)
    '*************************************************************
    pbrAbschluss.value = 1050
    frmWKL27!Label2(1).Caption = "Artikelgruppendaten werden übernommen..."
    frmWKL27!Label2(1).Refresh

    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    

    If NewTableSuchenDBKombi("AGN_IN", dbDb) Then
        lRet = fnVerarbeiteArtGruppenMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "AGN_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 9: ARTLief (ARTLiefKombinationen)
    '*************************************************************
    pbrAbschluss.value = 1100
    frmWKL27!Label2(1).Caption = "Artikel-Lieferantendaten werden übernommen..."
    frmWKL27!Label2(1).Refresh

    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    

    If NewTableSuchenDBKombi("ARTL_IN", dbDb) Then
        lRet = fnVerarbeiteArtliefMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ARTL_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 10: FIL_IN
    '*************************************************************
    pbrAbschluss.value = 1200
    frmWKL27!Label2(1).Caption = "Filialen werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("FIL_IN", dbDb) Then
        lRet = fnVerarbeiteFILIALENMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "FIL_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 11: ZUG_IN
    '*************************************************************
    pbrAbschluss.value = 1300
    frmWKL27!Label2(1).Caption = "Zugriffsrechte werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("ZUG_IN", dbDb) Then
        lRet = fnVerarbeiteBEDZUGRIMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ZUG_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 12: FARB_IN
    '*************************************************************
    pbrAbschluss.value = 1350
    frmWKL27!Label2(1).Caption = "Bestelldaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("ZBREST_IN", dbDb) Then
        lRet = fnVerarbeiteZBRESTMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ZBREST_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 12: ZUNTER_IN
    '*************************************************************
    
    pbrAbschluss.value = 1360
    frmWKL27!Label2(1).Caption = "Unterwegsdaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("ZUNTER_IN", dbDb) Then
        lRet = fnVerarbeiteZUNTERMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ZUNTER_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    
    '*************************************************************
    '* Schritt 12: ZBLOCK_IN
    '*************************************************************
    
    pbrAbschluss.value = 1370
    frmWKL27!Label2(1).Caption = "geblockte Artikel werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("ZBLOCK_IN", dbDb) Then
        lRet = fnVerarbeiteZBLOCKMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ZBLOCK_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 12: ZSPERR_IN
    '*************************************************************
    
    pbrAbschluss.value = 1380
    frmWKL27!Label2(1).Caption = "gesperrte Artikel werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("ZSPERR_IN", dbDb) Then
        lRet = fnVerarbeiteZSPERRMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "ZSPERR_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    '*************************************************************
    '* Schritt 12: Bonus Auszahlungen
    '*************************************************************
    
    pbrAbschluss.value = 1380
    frmWKL27!Label2(1).Caption = "Bonus Auszahlungen werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("KUNDENBONUS_IN", dbDb) Then
        lRet = fnVerarbeiteZBONUSRUECKMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "KUNDENBONUS_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 12: Bonus SYS Auszahlungen
    '*************************************************************
    
    pbrAbschluss.value = 1400
    frmWKL27!Label2(1).Caption = "Bonus SYS werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("BONUS_SYS_IN", dbDb) Then
        lRet = fnVerarbeiteZBONUS_SYS(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "BONUS_SYS_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    
    
    
    '*************************************************************
    '* Schritt 3a: BONUSL_IN.DBF (Bonus Auszahlungen löschen)
    '*************************************************************
    pbrAbschluss.value = 1420
    frmWKL27!Label2(1).Caption = "Bonus Auszahlungen werden gelöscht..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    

    If NewTableSuchenDBKombi("BONUSL_IN", dbDb) Then
        lRet = fnVerarbeiteBONUSL_INMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "BONUSL_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    
    
    
'

    '*************************************************************
    '* Schritt 3a: PICKLISTE_IN.DBF (PICKLISTE)
    '*************************************************************
    pbrAbschluss.value = 1450
    frmWKL27!Label2(1).Caption = "PICKLISTE wird verarbeitet..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    

    If NewTableSuchenDBKombi("PICKLISTE_IN", dbDb) Then
        lRet = fnVerarbeitePICKLISTE_INMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "PICKLISTE_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    
    
    '*************************************************************
    '* Schritt 3a: KREDITL_IN.DBF (geschriebene Kredite löschen)
    '*************************************************************
    pbrAbschluss.value = 1500
    frmWKL27!Label2(1).Caption = "geschriebene Kredite werden gelöscht..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    

    If NewTableSuchenDBKombi("KREDITL_IN", dbDb) Then
        lRet = fnVerarbeiteKREDITL_INMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "KREDITL_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    
    
    
    '*************************************************************
    '* Schritt 13: KUKASS_IN
    '*************************************************************
    pbrAbschluss.value = 1600
    frmWKL27!Label2(1).Caption = "Kundenverkaufsdaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("KUKASS_IN", dbDb) Then
        lRet = fnVerarbeiteKUKASSMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "KUKASS_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 14: PRSTERM_IN
    '*************************************************************
    pbrAbschluss.value = 1700
    frmWKL27!Label2(1).Caption = "Terminpreisdaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("PRSTERM_IN", dbDb) Then
        lRet = fnVerarbeitePRSTERMMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "PRSTERM_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 14: PREISTERM_IN
    '*************************************************************
    pbrAbschluss.value = 1800
    frmWKL27!Label2(1).Caption = "Terminpreisdaten werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("PREISTERM_IN", dbDb) Then
        lRet = fnVerarbeitePREISTERMMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "PREISTERM_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    '*************************************************************
    '* Schritt 15: BON_IN
    '*************************************************************
    pbrAbschluss.value = 1900
    frmWKL27!Label2(1).Caption = "Bontext werden synchronisiert..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "0"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Caption = "0"
    frmWKL27!Label2(5).Refresh
    
    If NewTableSuchenDBKombi("BON_IN", dbDb) Then
        lRet = fnVerarbeiteBonTextMOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "BON_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    '*************************************************************
    '* Schritt 15: Bestellungen_Export_IN.DBF
    '*************************************************************
    pbrAbschluss.value = 1950
    frmWKL27!Label2(1).Caption = "Bestellungen werden übernommen..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
 
    If NewTableSuchenDBKombi("Bestellungen_Export_IN", dbDb) Then
        lRet = fnVerarbeite_Bestellungen_Export_MOD6(dbDb)
        If lRet = 0 Then
            If iLoeschen = vbYes Then
                loeschNEW "Bestellungen_Export_IN", dbDb
            End If
        End If
    Else
    
    End If
    
    
    
    
    
    pbrAbschluss.value = 2000
    frmWKL27!Label2(1).Caption = "übertragene Filialdateien werden gelöscht..."
    frmWKL27!Label2(1).Refresh
    
    frmWKL27!Label2(3).Caption = "alle"
    frmWKL27!Label2(3).Refresh
    
    frmWKL27!Label2(5).Visible = False
    frmWKL27!Label2(4).Visible = False
    

    
    dbDb.Close
    
    pbrAbschluss.Visible = False
    
    SucheFilialKassenDateienMOD6 = True
  
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "SucheFilialKassenDateienMOD6"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function HatArtikelVerkäufe(lartnr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    HatArtikelVerkäufe = False
    
    cSQL = "Select * from UMS_ART  where ARTNR = " & lartnr
    cSQL = cSQL & " and Jahr = " & Year(DateValue(Now)) - 1
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        HatArtikelVerkäufe = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If HatArtikelVerkäufe = False Then
        cSQL = "Select * from UMS_ART  where ARTNR = " & lartnr
        cSQL = cSQL & " and Jahr = " & Year(DateValue(Now))
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            HatArtikelVerkäufe = True
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "HatArtikelVerkäufe"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Function

Public Function MBgleichNull(lartnr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    MBgleichNull = False
    
    cSQL = "Select * from Artikel where ARTNR = " & lartnr
    cSQL = cSQL & " and MINBEST = 0 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        MBgleichNull = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "MBgleichNull"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Function


