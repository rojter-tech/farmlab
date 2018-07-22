Private Sub CommandButton1_Click()
    Dim neat As Integer, spike As Integer
    Worksheet_Change (Cells(1, 2))
    CompInt = LoopRowsStoreNeat
    LoopRowsStoreSpike
    CreateCompoundSheets
    neat = CheckNeatIndex
    spike = CheckSpikeIndex
    TransferControlRange (CompInt)
    If Not neat = spike Then
        MsgBox "Antalet injektioner för neat och spike överensstämmer inte"
    End If
End Sub
Private Function LoopRowsStoreNeat() As Integer
'Går igenom rad för rad samtliga substanser i CompleteSummary och lagrar metadata för Neat i fliken CompoundList
    Dim j As Integer, CompRange As Range, Cell As Range, WShComp As Worksheet, InfoSheat As Worksheet
    Dim strSheetName As String, ActWSh As Worksheet, bln As Boolean
    Dim yourString, subString, replacementString, newString As String

    'Definera området med substansinformation
    Set InfoSheat = Sheets("Neat")
    Set CompRange = InfoSheat.Range(InfoSheat.Cells(1, 1), InfoSheat.Cells(Rows.Count, "A").End(xlUp))

    'Bekräfta att bladet CompoundsList inte redan existerar.
    strSheetName = Trim("CompoundsList")
    On Error Resume Next
    Set ActWSh = ActiveWorkbook.Worksheets(strSheetName)
    On Error Resume Next
    If Not ActWSh Is Nothing Then
        bln = True
    Else
        bln = False
        Err.Clear
    End If

    'Om bladet CompoundsList inte existerar, skapa det.
    If bln = False Then
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = "CompoundsList"
    End If

    'Fyll x första cellerna med Compound strängar i CompoundsList
    Set WShComp = Sheets("CompoundsList")
    j = 1
    WShComp.Cells(1, 1).Value = "Compound"
    WShComp.Cells(1, 2).Value = "Header Row Neat"
    For Each Cell In CompRange
        If Cell.Value Like "*Compound*" And Cell.Value Like "*:*" Then
            yourString = Cell.Value
            If j < 10 Then
                subString = Right(yourString, Len(yourString) - 13)
            ElseIf j < 100 Then
                subString = Right(yourString, Len(yourString) - 14)
            ElseIf j < 1000 Then
                subString = Right(yourString, Len(yourString) - 15)
            Else
                MsgBox "Du har klistrat in för många CompleteSummary i basfliken"
            End If
            replacementString = ""
            newString = replacementString + subString
            WShComp.Cells(j + 1, 1).Value = newString
            WShComp.Cells(j + 1, 2).Value = Cell.Row + 2
            j = j + 1
        Else
        End If
    Next
    LoopRowsStoreNeat = j - 1

End Function
Private Sub LoopRowsStoreSpike()
'Lagrar metadata för Spike i fliken CompoundList
    Dim j As Integer, CompRange As Range, Cell As Range, WShComp As Worksheet, InfoSheat As Worksheet
    Dim strSheetName As String, ActWSh As Worksheet, bln As Boolean

    'Definera området med substansinformation
    Set InfoSheat = Sheets("Spike")
    Set CompRange = InfoSheat.Range(InfoSheat.Cells(3, 1), InfoSheat.Cells(Rows.Count, "A").End(xlUp))

    'Fyll x första cellerna med Compound strängar i CompoundsList
    Set WShComp = Sheets("CompoundsList")
    j = 1
    WShComp.Cells(1, 3).Value = "Header Row Spike"
    For Each Cell In CompRange
        If Cell.Value Like "*Compound*" And Cell.Value Like "*:*" Then
            WShComp.Cells(j + 1, 3).Value = Cell.Row + 2
            j = j + 1
        Else
        End If
    Next

End Sub
Private Sub CreateCompoundSheets()
'Skapar flikar från metadata
    Dim CompRange As Range, Cell As Range, WShComp As Worksheet, SheetName As String
    Dim strSheetName As String, ActWSh As Worksheet, bln As Boolean
    Set WShComp = Sheets("CompoundsList")
    Set CompRange = WShComp.Range(WShComp.Cells(2, 1), WShComp.Cells(Rows.Count, "A").End(xlUp))

    For Each Cell In CompRange
        'Bekräfta att bladen från CompoundsList inte redan existerar.
        strSheetName = Trim(Cell.Value)
        On Error Resume Next
        Set ActWSh = ActiveWorkbook.Worksheets(strSheetName)
        On Error Resume Next
        If Not ActWSh Is Nothing Then
            bln = True
        Else
            bln = False
            Err.Clear
        End If

        If bln = False Then
            Sheets.Add After:=Sheets(Sheets.Count)
            Sheets(Sheets.Count).Name = Cell
        End If
    Next

End Sub
Private Function CheckNeatIndex() As Integer
'Kollar och lagrar indexdata för Neat
    Dim IndexRange As Range, InfoSheat As Worksheet, i As Integer, j As Integer, WShComp As Worksheet
    Set WShComp = Sheets("CompoundsList")
    Set InfoSheat = Sheets("Neat")
    Set IndexRange = InfoSheat.Range(InfoSheat.Cells(1, 1), InfoSheat.Cells(Rows.Count, "A").End(xlUp))
    i = 1
    j = 1
    Do
        i = i + 1
    Loop Until IsNumeric(InfoSheat.Cells(i, 1).Value) = True And IsEmpty(InfoSheat.Cells(i, 1).Value) = False
    Do
        j = j + 1
        i = i + 1
    Loop Until IsNumeric(InfoSheat.Cells(i, 1).Value) = False Or IsEmpty(InfoSheat.Cells(i, 1).Value) = True
    WShComp.Cells(1, 4).Value = "Neat InjectionNumber"
    WShComp.Cells(2, 4).Value = j - 1
    CheckNeatIndex = j - 1
End Function
Private Function CheckSpikeIndex() As Integer
'Kollar och lagrar indexdata för Spike
    Dim IndexRange As Range, InfoSheat As Worksheet, i As Integer, j As Integer, WShComp As Worksheet
    Set WShComp = Sheets("CompoundsList")
    Set InfoSheat = Sheets("Spike")
    Set IndexRange = InfoSheat.Range(InfoSheat.Cells(1, 1), InfoSheat.Cells(Rows.Count, "A").End(xlUp))
    i = 1
    j = 1
    Do
        i = i + 1
    Loop Until IsNumeric(InfoSheat.Cells(i, 1).Value) = True And IsEmpty(InfoSheat.Cells(i, 1).Value) = False
    Do
        j = j + 1
        i = i + 1
    Loop Until IsNumeric(InfoSheat.Cells(i, 1).Value) = False Or IsEmpty(InfoSheat.Cells(i, 1).Value) = True
    WShComp.Cells(1, 5).Value = "Spike InjectionNumber"
    WShComp.Cells(2, 5).Value = j - 1
    CheckSpikeIndex = j - 1
End Function
Private Sub TransferControlRange(ByVal CompInt As Integer)
    Dim HeaderRange As Range, ControlRange As Range, NeatSheet As Worksheet, SpikeSheet As Worksheet, WShComp As Worksheet
    Dim headerneat As Integer, headerspike As Integer, injectionsneat As Integer
    Dim IDRow As Integer, StdConcRow As Integer, RTRow As Integer, PredRTRow As Integer, AreaRow As Integer, RFRow As Integer, TypeRow As Integer
    Dim NeatArea As Long, SpikeArea As Long
    Dim j As Integer, k As Integer
    Set WShComp = Sheets("CompoundsList")
    Set NeatSheet = Sheets("Neat")
    Set SpikeSheet = Sheets("Spike")
    injectionsneat = WShComp.Cells(2, 4)

    For k = 1 To CompInt
        headerneat = WShComp.Cells(k + 1, 2)
        headerspike = WShComp.Cells(k + 1, 3)

        Set HeaderRange = NeatSheet.Range(NeatSheet.Cells(headerneat, 1), NeatSheet.Cells(headerneat, Columns.Count).End(xlToLeft))
        Set ControlRange = NeatSheet.Range(NeatSheet.Cells(headerneat, 3), NeatSheet.Cells(headerneat + injectionsneat, 3))
        For Each Cell In HeaderRange
            If Cell.Value = "ID" Then
                IDRow = CInt(Cell.Column)
            End If

            If Cell.Value Like "*Std*" And Cell.Value Like "*Conc*" Then
                StdConcRow = CInt(Cell.Column)
            End If

            If Cell.Value = "RT" Then
                RTRow = CInt(Cell.Column)
            End If

            If Cell.Value Like "*Pred*" And Cell.Value Like "*RT*" Then
                PredRTRow = CInt(Cell.Column)
            End If

            If Cell.Value = "Area" Then
                AreaRow = CInt(Cell.Column)
            End If

            If Cell.Value Like "*Ratio*" And Cell.Value Like "*Flag*" Then
                RFRow = CInt(Cell.Column)
            End If

            If Cell.Value = "Type" Then
                TypeRow = CInt(Cell.Column)
            End If
        Next

        For Each Cell In HeaderRange
            If Cell.Value = "Type" Then
                j = 1
                For i = 1 To injectionsneat
                    If NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "Standard" Or NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "QC" Then
                        NeatArea = NeatSheet.Cells(headerneat + i, AreaRow).Value
                        SpikeArea = SpikeSheet.Cells(headerspike + i, AreaRow).Value
                        Sheets(CStr(WShComp.Cells(k + 1, 1))).Cells(1, 1).Value = "Sample"
                        Sheets(CStr(WShComp.Cells(k + 1, 1))).Cells(j + 1, 1).Value = NeatSheet.Cells(headerneat + i, IDRow).Value
                        Sheets(CStr(WShComp.Cells(k + 1, 1))).Cells(1, 2).Value = "TAC"
                        Sheets(CStr(WShComp.Cells(k + 1, 1))).Cells(j + 1, 2).Value = NeatArea / (SpikeArea - NeatArea)
                        Sheets(CStr(WShComp.Cells(k + 1, 1))).Cells(1, 9).Value = "RatioFlag"
                        Sheets(CStr(WShComp.Cells(k + 1, 1))).Cells(j + 1, 9).Value = NeatSheet.Cells(headerneat + i, RFRow).Value
                        j = j + 1
                    End If
                Next
            End If
        Next
    Next

End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address <> "$A$1" Then Exit Sub
    If IsEmpty(Target) Then Exit Sub
    If Len(Target.Value) > 31 Then
        Application.EnableEvents = False
        Target.ClearContents
        Application.EnableEvents = True
        Exit Sub
    End If

    Dim IllegalCharacter(1 To 7) As String, i As Integer
    IllegalCharacter(1) = "/"
    IllegalCharacter(2) = "\"
    IllegalCharacter(3) = "["
    IllegalCharacter(4) = "]"
    IllegalCharacter(5) = "*"
    IllegalCharacter(6) = "?"
    IllegalCharacter(7) = ":"
    For i = 1 To 7
        If InStr(Target.Value, (IllegalCharacter(i))) > 0 Then
            Application.EnableEvents = False
            Target.ClearContents
            Application.EnableEvents = True
            Exit Sub
        End If
    Next i

    Dim strSheetName As String, ActWSh As Worksheet, bln As Boolean
    strSheetName = Trim(Target.Value)
    On Error Resume Next
    Set ActWSh = ActiveWorkbook.Worksheets(strSheetName)
    On Error Resume Next
    If Not ActWSh Is Nothing Then
        bln = True
    Else
        bln = False
        Err.Clear
    End If

    If bln = False Then
        ActiveSheet.Name = strSheetName
    Else
        Application.EnableEvents = False
        Target.ClearContents
        Application.EnableEvents = True
    End If

End Sub
