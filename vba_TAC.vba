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
'Går igenom rad för rad samtliga substanser i CompleteSummary och lagrar metadata för Neat i fliken MetaData
    Dim j As Integer, CompRange As Range, Cell As Range, MetaDataSh As Worksheet, InfoSheat As Worksheet
    Dim strSheetName As String, ActWSh As Worksheet, bln As Boolean
    Dim yourString, subString, replacementString, newString As String
    
    'Definera området med substansinformation
    Set InfoSheat = Sheets("Neat")
    Set CompRange = InfoSheat.Range(InfoSheat.Cells(1, 1), InfoSheat.Cells(Rows.Count, "A").End(xlUp))
    
    'Bekräfta att bladet MetaData inte redan existerar.
    strSheetName = Trim("MetaData")
    On Error Resume Next
    Set ActWSh = ActiveWorkbook.Worksheets(strSheetName)
    On Error Resume Next
    If Not ActWSh Is Nothing Then
        bln = True
    Else
        bln = False
        Err.Clear
    End If
    
    'Om bladet MetaData inte existerar, skapa det.
    If bln = False Then
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = "MetaData"
    End If
    
    'Fyll x första cellerna med Compound strängar i MetaData
    Set MetaDataSh = Sheets("MetaData")
    j = 1
    MetaDataSh.Cells(1, 1).Value = "Compound"
    MetaDataSh.Cells(1, 2).Value = "Header Row Neat"
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
            MetaDataSh.Cells(j + 1, 1).Value = newString
            MetaDataSh.Cells(j + 1, 2).Value = Cell.Row + 2
            j = j + 1
        Else
        End If
    Next
    LoopRowsStoreNeat = j - 1

End Function
Private Sub LoopRowsStoreSpike()
'Lagrar metadata för Spike i fliken CompoundList
    Dim j As Integer, CompRange As Range, Cell As Range, MetaDataSh As Worksheet, InfoSheat As Worksheet
    Dim strSheetName As String, ActWSh As Worksheet, bln As Boolean
    
    'Definera området med substansinformation
    Set InfoSheat = Sheets("Spike")
    Set CompRange = InfoSheat.Range(InfoSheat.Cells(3, 1), InfoSheat.Cells(Rows.Count, "A").End(xlUp))
    
    'Fyll x första cellerna med Compound strängar i MetaData
    Set MetaDataSh = Sheets("MetaData")
    j = 1
    MetaDataSh.Cells(1, 3).Value = "Header Row Spike"
    For Each Cell In CompRange
        If Cell.Value Like "*Compound*" And Cell.Value Like "*:*" Then
            MetaDataSh.Cells(j + 1, 3).Value = Cell.Row + 2
            j = j + 1
        Else
        End If
    Next

End Sub
Private Sub CreateCompoundSheets()
'Skapar flikar från metadata
    Dim CompRange As Range, Cell As Range, MetaDataSh As Worksheet, SheetName As String
    Dim strSheetName As String, ActWSh As Worksheet, bln As Boolean
    Set MetaDataSh = Sheets("MetaData")
    Set CompRange = MetaDataSh.Range(MetaDataSh.Cells(2, 1), MetaDataSh.Cells(Rows.Count, "A").End(xlUp))
    
    For Each Cell In CompRange
        'Bekräfta att bladen från MetaData inte redan existerar.
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
    Dim IndexRange As Range, InfoSheat As Worksheet, i As Integer, j As Integer, MetaDataSh As Worksheet
    Set MetaDataSh = Sheets("MetaData")
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
    MetaDataSh.Cells(1, 4).Value = "Neat InjectionNumber"
    MetaDataSh.Cells(2, 4).Value = j - 1
    CheckNeatIndex = j - 1
End Function
Private Function CheckSpikeIndex() As Integer
'Kollar och lagrar indexdata för Spike
    Dim IndexRange As Range, InfoSheat As Worksheet, i As Integer, j As Integer, MetaDataSh As Worksheet
    Set MetaDataSh = Sheets("MetaData")
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
    MetaDataSh.Cells(1, 5).Value = "Spike InjectionNumber"
    MetaDataSh.Cells(2, 5).Value = j - 1
    CheckSpikeIndex = j - 1
End Function
Private Sub TransferControlRange(ByVal CompInt As Integer)
    Dim HeaderRange As Range, ControlRange As Range, NeatSheet As Worksheet, SpikeSheet As Worksheet, MetaDataSh As Worksheet, CompSheet As Worksheet
    Dim headerneat As Integer, headerspike As Integer, injectionsneat As Integer
    Dim IDRow As Integer, StdConcRow As Integer, RTRow As Integer, PredRTRow As Integer, AreaRow As Integer, RFRow As Integer, TypeRow As Integer
    Dim NeatArea As Long, SpikeArea As Long
    Dim j As Integer, k As Integer, kal As Integer, compName As String
    Set MetaDataSh = Sheets("MetaData")
    Set NeatSheet = Sheets("Neat")
    Set SpikeSheet = Sheets("Spike")
    injectionsneat = MetaDataSh.Cells(2, 4)
    MetaDataSh.Cells(1, 6).Value = "Calibration Points"
    
    For k = 1 To CompInt
        headerneat = MetaDataSh.Cells(k + 1, 2)
        headerspike = MetaDataSh.Cells(k + 1, 3)
        compName = CStr(MetaDataSh.Cells(k + 1, 1))
        Set CompSheet = ThisWorkbook.Sheets(compName)
        CompSheet.Cells(1, 1).Value = "Sample"
        CompSheet.Cells(1, 2).Value = "TAC"
        CompSheet.Cells(1, 3).Value = "Conc"
        CompSheet.Cells(1, 9).Value = "RatioFlag"
        
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
                kal = 0
                For i = 1 To injectionsneat
                    If NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "Standard" Or NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "QC" Then
                        NeatArea = NeatSheet.Cells(headerneat + i, AreaRow).Value
                        SpikeArea = SpikeSheet.Cells(headerspike + i, AreaRow).Value
                        CompSheet.Cells(j + 1, 1).Value = NeatSheet.Cells(headerneat + i, IDRow).Value
                        CompSheet.Cells(j + 1, 2).Value = NeatArea / (SpikeArea - NeatArea)
                        CompSheet.Cells(j + 1, 9).Value = NeatSheet.Cells(headerneat + i, RFRow).Value
                        j = j + 1
                    End If
                    If NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "Standard" Then
                    MetaDataSh.Cells(i - 2, 6).Value = NeatSheet.Cells(headerneat + i, StdConcRow).Value
                    kal = kal + 1
                    End If
                Next
                j = 1
                CalExecutor kal, compName
                For i = 1 To injectionsneat
                    If NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "Standard" Or NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "QC" Then
                        CompSheet.Cells(j + 1, 3).Value = (CompSheet.Cells(j + 1, 2).Value - CompSheet.Range("K2")) / (CompSheet.Range("J2"))
                        j = j + 1
                    End If
                Next
            End If
        Next
    Next
End Sub
Private Sub CalExecutor(ByVal CompInt As Integer, ByVal compName As String)
    Dim MetaDataSh As Worksheet, CompSheet As Worksheet, Chrt As Chart, yL As Range, xL As Range
    Set CompSheet = ThisWorkbook.Sheets(compName)
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    Set Chrt = CompSheet.Shapes.AddChart2.Chart
    Set yL = CompSheet.Range("$B$2:$B$7")
    Set xL = MetaDataSh.Range("$F$2:$F$7")
    SlopeValue = Application.WorksheetFunction.Slope(yL, xL)
    InterceptValue = Application.WorksheetFunction.Intercept(yL, xL)
    
    With Chrt
        .Parent.Left = CompSheet.Range("K4").Left
        .Parent.Top = CompSheet.Range("K4").Top
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        .ChartType = xlXYScatter
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Calibrationcurve: " & compName
            .Values = "='" & compName & "'!$B$2:$B$7"
            .XValues = "='MetaData'!$F$2:$F$7"
            .Trendlines.Add
            .Trendlines(1).Select
        End With
        Selection.DisplayEquation = True
        Selection.DisplayRSquared = True
        .HasTitle = True
        .ChartTitle.Characters.Text = "Calibrationcurve: " & compName
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Concentration [ng/mL]"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "TAC Ratio"
        .Axes(xlCategory).HasMajorGridlines = True
        .Axes(xlCategory).HasMinorGridlines = False
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).HasMinorGridlines = True
        .HasLegend = False
    End With
    CompSheet.Range("J1").Value = "Slope"
    CompSheet.Range("J2").Value = SlopeValue
    CompSheet.Range("K1").Value = "Intercept"
    CompSheet.Range("K2").Value = InterceptValue
    
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
