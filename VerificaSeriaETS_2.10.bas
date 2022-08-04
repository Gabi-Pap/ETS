Attribute VB_Name = "Module1"
'Verificare Seria ETS Versiune: 2.8'
'Data: 21 Iunie 2022'

Sub VerificareSerie()
Dim FindString As String
    Dim Rng As Range
    Dim poz As Integer
    Dim colName As String
    Dim colNameIndex As Integer
    Dim colNameOld As String
    Dim colNumber As Integer
    Dim fisa As String
    Dim lastColumn As Long
    Dim CopyRange As String
    Dim an As String
    Dim anInclus As Boolean
    Dim anInitial As String
    Dim colSerieCitita As Integer
    Dim colAnFabricatieInitial As Integer
    Dim colCautare As Integer
    Dim SArray() As String
    Dim Temp As String
    Dim answ As Integer
    
    FindString = " "
    
    fisa = InputBox("Seriile care se cauta se gasesc pe fisa", , "Sheet1")
    If StrPtr(fisa) = 0 Then
        Exit Sub
    End If
    
    'colName = InputBox("Seriile care se cauta se gasesc pe coloana", , "G")'
    'If StrPtr(colName) = 0 Then'
    '    Exit Sub'
    'End If'
    
    an = InputBox("Anul seriilor (0: nu scrie anul)", , "0")
    If StrPtr(an) = 0 Then
        Exit Sub
    End If
    
    'colNumber = Range(colName & 1).Column'
    colSerieCitita = 0
    colAnFabricatieInitial = 0
    lastColumn = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    For i = 1 To lastColumn
        If Cells(1, i).Value = "Serie citita" Then colSerieCitita = i
        If Cells(1, i).Value = "An fabricatie initial" Then colAnFabricatieInitial = i
        If Cells(1, i).Value = "Serii corectate" Then colCautare = i
        If Cells(1, i).Value = "Serie Producator" Then colNameIndex = i
    Next i
    colName = Split(Cells(, colNameIndex).Address, "$")(1) 'conversie numar coloana in nume coloana
    colNumber = Range(colName & 1).Column
    If colSerieCitita = 0 Then
        colSerieCitita = lastColumn + 1
        colAnFabricatieInitial = lastColumn + 2
        colCautare = lastColumn + 3
        Cells(1, colSerieCitita) = "Serie citita"
        Cells(1, colAnFabricatieInitial) = "An fabricatie initial"
        Cells(1, colCautare) = "Serii corectate"
    End If
    If colSerieCitita > 0 And colCautare = 0 Then colCautare = lastColumn + 1
    If Cells(2, colCautare) = vbNullString Then
        Cells(1, colCautare) = "Serii corectate"
        For i = 2 To ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, colCautare).NumberFormat = "@"
            Cells(i, colCautare) = Cells(i, colName)
            Cells(i, colCautare) = Replace(Cells(i, colCautare), "#", "")
            Cells(i, colCautare) = Replace(Cells(i, colCautare), "*", "")
        Next i
    End If
    colNameOld = colName
    colName = Split(Cells(1, colCautare).Address, "$")(1) 'numele coloanei unde sunt seriile corectate
    Let CopyRange = colName & Startrow & ":" & colName & Lastrow
    While FindString <> vbNullString
    FindString = InputBox("Citeste seria")
    If StrPtr(FindString) = 0 Then
        Exit Sub
    End If
    If FindString = vbNullString Then
        answ = MsgBox("Continuati cu un alt an?", vbOKCancel)
        If answ = 2 Then
            Exit Sub
        Else
            an = InputBox("Anul seriilor (0: nu scrie anul)", , "0")
            If StrPtr(an) = 0 Then 'apasa Cancel
                Exit Sub
            End If
            FindString = InputBox("Citeste seria")
            If StrPtr(FindString) = 0 Then 'apasa Cancel
                Exit Sub
            End If
            If FindString = vbNullString Then 'nu scrie nimic
                Exit Sub
            End If
        End If
    End If
    
    FindString = Replace(FindString, "|", "")
    
    'electromagnetica?
    SArray = Split(FindString, " ") 'separatorul este blanc
    If UBound(SArray) = 2 Then 'contor electromagnetica
        FindString = SArray(2) + "/20" + SArray(1)
    ElseIf UBound(SArray) = 1 Then
        FindString = SArray(0) + "/" + SArray(1)
    'Else
    '    FindString = SArray(0) 'contor normal
    End If
    
    Temp = "" 'eliminare caractere in fata (OCG, QCG)
    For k = 1 To Len(FindString)
        If (IsNumeric(Mid(FindString, k, 1))) = True Then
            Temp = Temp & Mid(FindString, k, 1)
        End If
        If (Mid(FindString, k, 1) = "/") Then
            Temp = Temp & Mid(FindString, k, 1)
        End If
    Next k
    FindString = Temp
    
    poz = InStr(FindString, "/")
    If Trim(FindString) <> "" Then
        If poz <> 0 And UBound(SArray) = 2 Then
            Serie = Val(Mid(FindString, 1, poz - 1))
        ElseIf poz <> 0 Then
            Serie = Mid(FindString, 1, poz - 1)
        Else
            Serie = FindString
        End If
        With Sheets(fisa).Range(CopyRange)
            Set Rng = .Find(What:=Serie, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlPart, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False) 'cautare normala
            If (Rng Is Nothing) And (Len(Serie) > 7) Then 'seria trebuie sa aiba cel prin 4 caractere
                Set Rng = .Find(What:=Mid(Serie, 1, Len(Serie) - 4), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False) 'cautare serie cu an inclus
                If Not Rng Is Nothing Then anInclus = True Else anInclus = False
            End If
            If (Rng Is Nothing) And (Mid(Serie, 1, 1) = "0") Then
                Set Rng = .Find(What:=Mid(Serie, 2, Len(Serie)), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False) 'cautare serie care are un 0 in fata
            End If
            If (Rng Is Nothing) And (Mid(Serie, 1, 2) = "00") Then
                Set Rng = .Find(What:=Mid(Serie, 3, Len(Serie)), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False) 'cautare serie care are un 00 in fata
            End If
            If Not Rng Is Nothing Then
                Application.GoTo Rng, False
                Rng.Interior.ColorIndex = 37
                Cells(Rng.Row, colNameOld).Interior.ColorIndex = 37
                Cells(Rng.Row, colSerieCitita).NumberFormat = "@"
                anInitial = Cells(Rng.Row, "H")
                    
                If poz > 0 Then 'are slash
                    Cells(Rng.Row, colSerieCitita) = FindString
                    If Len(Mid(FindString, poz + 1, Len(FindString))) < 4 Then 'anul are 4 cifre
                        If Mid(FindString, poz + 1, Len(FindString)) > 20 Then
                            Cells(Rng.Row, "H") = "19" + Mid(FindString, poz + 1, Len(FindString))
                            If anInitial <> Cells(Rng.Row, "H") Then Cells(Rng.Row, colAnFabricatieInitial) = anInitial
                        Else 'an 2000
                            Cells(Rng.Row, "H") = "20" + Mid(FindString, poz + 1, Len(FindString))
                            If anInitial <> Cells(Rng.Row, "H") Then Cells(Rng.Row, colAnFabricatieInitial) = anInitial
                        End If
                    Else
                        Cells(Rng.Row, "H") = Mid(FindString, poz + 1, Len(FindString))
                        If anInitial <> Cells(Rng.Row, "H") Then Cells(Rng.Row, colAnFabricatieInitial) = anInitial
                    End If
                ElseIf (poz = 0 And an = "0") Then 'are an inclus
                    If (Mid(FindString, Len(FindString) - 3, Len(FindString)) > 2030) Or (Mid(FindString, Len(FindString) - 3, Len(FindString)) < 1960) Then
                        Rng.Interior.ColorIndex = 0
                        MsgBox "Serie: " + FindString + " an eronat: " + Mid(FindString, Len(FindString) - 3, Len(FindString))
                        GoTo NextIteration
                    End If
                    Cells(Rng.Row, colSerieCitita) = FindString
                    Cells(Rng.Row, "H") = Mid(FindString, Len(FindString) - 3, Len(FindString))
                    If anInitial <> Cells(Rng.Row, "H") Then Cells(Rng.Row, colAnFabricatieInitial) = anInitial
                Else 'fara slash si an <> 0
                    Cells(Rng.Row, colSerieCitita) = FindString + "/" + an
                    If Len(an) = 2 Then 'anul are doar doua cifre
                        If CInt(an) < 50 Then
                            Cells(Rng.Row, "H") = "20" + an 'an de 2 cifre completat la 4 in An fabricatie
                        Else
                            Cells(Rng.Row, "H") = "19" + an
                        End If
                    End If
                    'Cells(Rng.Row, "H") = an
                    If anInitial <> an Then
                        Cells(Rng.Row, colAnFabricatieInitial) = anInitial
                        Cells(Rng.Row, "H") = an
                    End If
                End If
                    
            Else
                MsgBox "Nu exista " + FindString
            End If
        End With
    End If
NextIteration:
    Wend
End Sub
