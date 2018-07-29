Private Sub EvaluateButton_Click()
    Worksheet_Change (Cells(1, 1))
    Dim neat As Integer, spike As Integer, c As Integer
    Dim StartRowNeat As Integer, StartColumnNeat As Integer, CompCtrl As Integer, CompColumn As Integer, NeatRange As Range, SpikeRange As Range
    Dim StartRowSpike As Integer, StartColumnSpike As Integer, SpikeSheet As Worksheet, NeatSheet As Worksheet
    Dim strSheetName As String, bln As Boolean, CtrlWSh As Worksheet
    Dim controlMap As Integer
    
    'Kollar om bladet MetaData existerar.
    strSheetName = Trim("MetaData")
    On Error Resume Next
    Set CtrlWSh = ThisWorkbook.Worksheets(strSheetName)
    On Error Resume Next
    If Not CtrlWSh Is Nothing Then
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
    
    ' Sätt upp Neat och Spike bladen
    Set NeatSheet = ThisWorkbook.Sheets("Neat")
    Set SpikeSheet = ThisWorkbook.Sheets("Spike")
    Set NeatRange = ThisWorkbook.Sheets("Neat").Range("A1:AZ500")
    Set SpikeRange = ThisWorkbook.Sheets("Spike").Range("A1:AZ500")
    'Letar efter position för inklistring av CompleteSummary i Neat
    For Each Cell In NeatRange
        If Cell.Value Like "*Compound*" Then
            StartRowNeat = CInt(Cell.Row)
            StartColumnNeat = CInt(Cell.Column)
            Exit For
        End If
    Next
    If StartRowNeat = 0 Then
        MsgBox "Det verkar inte finnas någon CompleteSummary inklistrat i Neat-bladet! Kopiera från TargetLynx"
        Exit Sub
    End If
    'Letar efter position för inklistring av CompleteSummary i Spike
    For Each Cell In SpikeRange
        If Cell.Value Like "*Compound*" Then
            StartRowSpike = CInt(Cell.Row)
            StartColumnSpike = CInt(Cell.Column)
            Exit For
        End If
    Next
    'Kontroll att värden finns
    If StartRowSpike = 0 Then
        MsgBox "Det verkar inte finnas någon CompleteSummary inklistrat i Spike-bladet! Kopiera från TargetLynx"
        Exit Sub
    End If
    compInt = LoopRowsStoreNeat(ByVal StartRowNeat, ByVal StartColumnNeat)
    MetaDataLayout compInt
    LoopRowsStoreSpike StartRowSpike, StartColumnSpike
    CreateCompoundSheets
    neat = CheckNeatIndex(ByVal StartRowNeat, ByVal StartColumnNeat)
    spike = CheckSpikeIndex(ByVal StartRowSpike, ByVal StartColumnSpike)
    ' Kontroll att antalet rader överensstämmer mellan Neat och Spike
    If Not neat = spike Then
        MsgBox "Antalet injektioner för neat och spike överensstämmer inte! Processa om från MassLynx till TargetLynx och klistra in igen."
        Exit Sub
    End If
    controlMap = TransferControlRange(ByVal compInt, ByVal StartRowNeat, ByVal StartColumnNeat, ByVal StartRowSpike, ByVal StartColumnSpike)
    'Kontroll att varje enskild SampleID översensstämmer mellan Neat och Spike
    If controlMap = 0 Then
        MsgBox "Beräkningen lyckades! Använd respektive substansflik för att utvärdera resulatet."
    Else
        MsgBox "Mappningen misslyckades, kontrollera att rader överensstämmer mellan Neat och Spike"
        Exit Sub
    End If
End Sub
Private Sub MetaDataLayout(ByVal compInt)
    Dim MetaDataRange As Range, MetaDataSh As Worksheet
    
    Set MetaDataRange = ThisWorkbook.Sheets("MetaData").Range("A1:F" & CStr(compInt + 1))
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    MetaDataSh.Cells(1, 1).Value = "Compound"
    MetaDataSh.Cells(1, 2).Value = "Header Row Neat"
    MetaDataSh.Cells(1, 3).Value = "Header Row Spike"
    MetaDataSh.Cells(1, 4).Value = "Neat InjectionNumber"
    MetaDataSh.Cells(1, 5).Value = "Spike InjectionNumber"
    MetaDataSh.Cells(1, 6).Value = "Calibration Points"
    MetaDataRange.EntireColumn.AutoFit
    With MetaDataRange
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlCenter
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = ColIndex
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = ColIndex
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = ColIndex
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = ColIndex
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = ColIndex
        End With
    End With
End Sub
Private Function LoopRowsStoreNeat(ByVal StartRow As Integer, ByVal StartColumn As Integer) As Integer
'Går igenom rad för rad samtliga substanser i CompleteSummary och lagrar metadata för Neat i fliken MetaData för att använda som referensvärden
    Dim j As Integer, CompRange As Range, Cell As Range, MetaDataSh As Worksheet, NeatSheet As Worksheet, CtrlWSh As Worksheet
    Dim yourString, subString, replacementString, newString As String

    'Definerar området med substansinformation
    Set NeatSheet = ThisWorkbook.Sheets("Neat")
    Set CompRange = NeatSheet.Range(NeatSheet.Cells(StartRow, StartColumn), NeatSheet.Cells(Rows.Count, StartColumn).End(xlUp))
    
    'Fyll x första cellerna med Compound strängar i MetaData
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    j = 1

    'Skalar av den del av strängen som innehåller "Compound #:"
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
Private Sub LoopRowsStoreSpike(ByVal StartRow As Integer, ByVal StartColumn As Integer)
'Lagrar metadata för Spike i bladet MetaData
    Dim j As Integer, CompRange As Range, Cell As Range, MetaDataSh As Worksheet, SpikeSheat As Worksheet
    Dim strSheetName As String, CtrlWSh As Worksheet, bln As Boolean
    
    'Definerar området med substansinformation
    Set SpikeSheat = ThisWorkbook.Sheets("Spike")
    Set CompRange = SpikeSheat.Range(SpikeSheat.Cells(StartRow, StartColumn), SpikeSheat.Cells(Rows.Count, StartColumn).End(xlUp))
    
    'Fyller x första cellerna med Compound-strängar i MetaData
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    j = 1
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
    Dim strSheetName As String, CtrlWSh As Worksheet, bln As Boolean
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    Set CompRange = MetaDataSh.Range(MetaDataSh.Cells(2, 1), MetaDataSh.Cells(Rows.Count, "A").End(xlUp))
    
    For Each Cell In CompRange
        'Bekräfta att bladen från MetaData inte redan existerar.
        strSheetName = Trim(Cell.Value)
        On Error Resume Next
        Set CtrlWSh = ThisWorkbook.Worksheets(strSheetName)
        On Error Resume Next
        If Not CtrlWSh Is Nothing Then
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
Private Function CheckNeatIndex(ByVal StartRow, ByVal StartColumn) As Integer
'Kollar och lagrar indexdata för Neat i MetaData
    Dim IndexRange As Range, InfoSheet As Worksheet, i As Integer, j As Integer, MetaDataSh As Worksheet
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    Set InfoSheet = ThisWorkbook.Sheets("Neat")
    Set IndexRange = InfoSheet.Range(InfoSheet.Cells(StartRow, StartColumn), InfoSheet.Cells(Rows.Count, StartColumn).End(xlUp))
    i = 1
    j = 1
    Do
        i = i + 1
    Loop Until IsNumeric(InfoSheet.Cells(StartRow - 1 + i, StartColumn).Value) = True And IsEmpty(InfoSheet.Cells(StartRow - 1 + i, StartColumn).Value) = False
    Do
        j = j + 1
        i = i + 1
    Loop Until IsNumeric(InfoSheet.Cells(StartRow - 1 + i, StartColumn).Value) = False Or IsEmpty(InfoSheet.Cells(StartRow - 1 + i, StartColumn).Value) = True
    MetaDataSh.Cells(2, 4).Value = j - 1
    CheckNeatIndex = j - 1
End Function
Private Function CheckSpikeIndex(ByVal StartRow, ByVal StartColumn) As Integer
'Kollar och lagrar indexdata för Spike i MetaData
    Dim IndexRange As Range, InfoSheet As Worksheet, i As Integer, j As Integer, MetaDataSh As Worksheet
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    Set InfoSheet = ThisWorkbook.Sheets("Spike")
    Set IndexRange = InfoSheet.Range(InfoSheet.Cells(StartRow, StartColumn), InfoSheet.Cells(Rows.Count, StartColumn).End(xlUp))
    i = 1
    j = 1
    Do
        i = i + 1
    Loop Until IsNumeric(InfoSheet.Cells(StartRow - 1 + i, StartColumn).Value) = True And IsEmpty(InfoSheet.Cells(StartRow - 1 + i, StartColumn).Value) = False
    Do
        j = j + 1
        i = i + 1
    Loop Until IsNumeric(InfoSheet.Cells(StartRow - 1 + i, StartColumn).Value) = False Or IsEmpty(InfoSheet.Cells(StartRow - 1 + i, StartColumn).Value) = True
    MetaDataSh.Cells(2, 5).Value = j - 1
    CheckSpikeIndex = j - 1
End Function

Private Function TransferControlRange(ByVal compInt As Integer, ByVal StartRowNeat As Integer, ByVal StartColumnNeat As Integer, ByVal StartRowSpike As Integer, ByVal StartColumnSpike As Integer) As Integer
    Dim HeaderRangeNeat As Range, ControlRange As Range, NeatSheet As Worksheet, SpikeSheet As Worksheet, MetaDataSh As Worksheet, CompSheet As Worksheet
    Dim headerneat As Integer, headerspike As Integer, injectionsneat As Integer
    Dim IDColumn As Integer, StdConcColumn As Integer, RTColumn As Integer, PredRTColumn As Integer, AreaColumnNeat As Integer, RFColumn As Integer, TypeColumn As Integer, AreaColumnSpike As Integer
    Dim NeatArea As Long, SpikeArea As Long, StdConcValue As Long, AvgRTValue As Double
    Dim a As Integer, j As Integer, k As Integer, m As Integer, compSheetPos As Integer, LastContCell As Integer, kal As Integer, compName As String, ConcValue As Long, ColIndex As Integer
    Dim IDColumnSpike As Integer, RFColumnSpike As Integer
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    Set NeatSheet = ThisWorkbook.Sheets("Neat")
    Set SpikeSheet = ThisWorkbook.Sheets("Spike")
    LastCal = CStr(compInt + compSheetPos)
    injectionsneat = MetaDataSh.Cells(2, 4)
    m = 1
    compSheetPos = 2
    controlMap = 0
    For k = 1 To compInt
        headerneat = MetaDataSh.Cells(k + 1, 2)
        headerspike = MetaDataSh.Cells(k + 1, 3)
        compName = CStr(MetaDataSh.Cells(k + 1, 1))
        Set CompSheet = ThisWorkbook.Sheets(compName)
        CompSheet.Cells(1, 2).Value = compName
        CompSheet.Cells(1, 2).Font.Bold = True
        CompSheet.Cells(1, 2).Font.Size = 15
        CompSheet.Cells(compSheetPos, 1).Value = "Sample" & Chr(10) & ""
        CompSheet.Cells(compSheetPos, 2).Value = "TAC" & Chr(10) & "Ratio"
        CompSheet.Cells(compSheetPos, 3).Value = "Conc" & Chr(10) & "ng/mL"
        CompSheet.Cells(compSheetPos, 4).Value = "Range" & Chr(10) & "(±20%)"
        CompSheet.Cells(compSheetPos, 5).Value = "Conc" & Chr(10) & "Criteria"
        CompSheet.Cells(compSheetPos, 6).Value = "RT" & Chr(10) & "Criteria"
        CompSheet.Cells(compSheetPos, 7).Value = "Injection" & Chr(10) & "Recovery"
        CompSheet.Cells(compSheetPos, 8).Value = "Injection" & Chr(10) & "Criteria"
        CompSheet.Cells(compSheetPos, 9).Value = "Ion Ratio" & Chr(10) & "Failed"
        CompSheet.Cells(compSheetPos, 10).Value = "Ion Ratio" & Chr(10) & "Criteria"
        CompSheet.Range("A1:I1").HorizontalAlignment = xlCenter
        
        Set HeaderRangeNeat = NeatSheet.Range(NeatSheet.Cells(headerneat, StartColumnNeat), NeatSheet.Cells(headerneat, Columns.Count).End(xlToLeft))
        Set HeaderRangeSpike = SpikeSheet.Range(SpikeSheet.Cells(headerspike, StartColumnSpike), SpikeSheet.Cells(headerspike, Columns.Count).End(xlToLeft))
        Set ControlRange = NeatSheet.Range(NeatSheet.Cells(headerneat, 3), NeatSheet.Cells(headerneat + injectionsneat, 3))
        For Each Cell In HeaderRangeNeat
            If Cell.Value = "ID" Then
                IDColumn = CInt(Cell.Column)
            End If
        
            If Cell.Value Like "*Std*" And Cell.Value Like "*Conc*" Then
                StdConcColumn = CInt(Cell.Column)
            End If
        
            If Cell.Value = "RT" Then
                RTColumn = CInt(Cell.Column)
            End If
        
            If Cell.Value Like "*Pred*" And Cell.Value Like "*RT*" Then
                PredRTColumn = CInt(Cell.Column)
            End If
        
            If Cell.Value = "Area" Then
                AreaColumnNeat = CInt(Cell.Column)
            End If
        
            If Cell.Value Like "*Ratio*" And Cell.Value Like "*Flag*" Then
                RFColumn = CInt(Cell.Column)
            End If
        
            If Cell.Value = "Type" Then
                TypeColumn = CInt(Cell.Column)
            End If
        Next
        
        For Each Cell In HeaderRangeSpike
            If Cell.Value = "Area" Then
                AreaColumnSpike = CInt(Cell.Column)
            End If
            If Cell.Value Like "*Pred*" And Cell.Value Like "*RT*" Then
                PredRTColumnSpike = CInt(Cell.Column)
            End If
            If Cell.Value = "ID" Then
                IDColumnSpike = CInt(Cell.Column)
            End If
            If Cell.Value Like "*Ratio*" And Cell.Value Like "*Flag*" Then
                RFColumnSpike = CInt(Cell.Column)
            End If
        Next
        
        AvgRTValue = (NeatSheet.Cells(headerneat + 1, PredRTColumn).Value + SpikeSheet.Cells(headerspike + 1, PredRTColumnSpike).Value) / 2
        CompSheet.Cells(1, 7).Value = "Average RT = " & CStr(AvgRTValue) & " min"
        
        For Each Cell In HeaderRangeNeat
            If Cell.Value = "Type" Then
                j = compSheetPos + 1
                kal = 0
                
                For i = 1 To injectionsneat
                    'TAC-Ratio för QC och Standard
                    If NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "Standard" Or NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "QC" Then
                        NeatArea = NeatSheet.Cells(headerneat + i, AreaColumnNeat).Value
                        SpikeArea = SpikeSheet.Cells(headerspike + i, AreaColumnSpike).Value
                        If NeatSheet.Cells(headerneat + i, IDColumn).Value = SpikeSheet.Cells(headerspike + i, IDColumnSpike).Value Then
                            CompSheet.Cells(j, 1).Value = NeatSheet.Cells(headerneat + i, IDColumn).Value
                        Else
                            TransferControlRange = 1
                            Exit Function
                        End If
                        
                        CompSheet.Cells(j, 2).Value = NeatArea / (SpikeArea - NeatArea)
                        CompSheet.Cells(j, 9).Value = NeatSheet.Cells(headerneat + i, RFColumn).Value & Chr(10) & ""
                        j = j + 1
                    End If
                    If NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "Standard" Then
                        m = m + 1
                        kal = kal + 1
                        MetaDataSh.Cells(m, 6).Value = NeatSheet.Cells(headerneat + i, StdConcColumn).Value
                    End If
                Next
                LastContCell = j - 1
                a = LastContCell + 4
                m = 1
                j = compSheetPos + 1
                
                'Kalibrera och skapa kontrolldata
                CalExecutor kal, compName, compSheetPos, LastContCell
                
                'Överför QC och Standard
                For i = 1 To injectionsneat
                    If NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "Standard" Or NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "QC" Then
                        ConcValue = (CompSheet.Cells(j, 2).Value - CompSheet.Range("L2")) / (CompSheet.Range("K2"))
                        CompSheet.Cells(j, 3).Value = ConcValue
                        StdConcValue = NeatSheet.Cells(headerneat + i, StdConcColumn).Value
                        CompSheet.Cells(j, 4).Value = CStr(StdConcValue - StdConcValue * 0.2) & " - " & CStr(StdConcValue + StdConcValue * 0.2)
                        If CompSheet.Cells(j, 3).Value > StdConcValue - StdConcValue * 0.2 And CompSheet.Cells(j, 3).Value < StdConcValue + StdConcValue * 0.2 Then
                            CompSheet.Cells(j, 5).Value = "PASS"
                            CompSheet.Cells(j, 5).Font.Bold = True
                            CompSheet.Cells(j, 5).Font.Color = RGB(Red:=2, Green:=147, Blue:=70)
                        Else
                            CompSheet.Cells(j, 5).Value = "FAIL"
                            CompSheet.Cells(j, 5).Font.Bold = True
                            CompSheet.Cells(j, 5).Font.Color = RGB(Red:=195, Green:=0, Blue:=4)
                        End If
                        If NeatSheet.Cells(headerneat + i, PredRTColumn).Value > AvgRTValue - 0.1 And NeatSheet.Cells(headerneat + i, PredRTColumn).Value < AvgRTValue + 0.1 Then
                            CompSheet.Cells(j, 6).Value = "PASS"
                            CompSheet.Cells(j, 6).Font.Bold = True
                            CompSheet.Cells(j, 6).Font.Color = RGB(Red:=2, Green:=147, Blue:=70)
                        Else
                            CompSheet.Cells(j, 6).Value = "FAIL"
                            CompSheet.Cells(j, 6).Font.Bold = True
                            CompSheet.Cells(j, 6).Font.Color = RGB(Red:=195, Green:=0, Blue:=4)
                        End If
                        j = j + 1
                    End If
                Next
                
                'Överför Analyte
                CompSheet.Cells(a - 1, 1).Value = "Sample"
                CompSheet.Cells(a - 1, 2).Value = "TAC"
                CompSheet.Cells(a - 1, 3).Value = "Conc"
                CompSheet.Cells(a - 1, 4).Value = "Ratio Criteria"
                c = compSheetPos + 1
                For i = 1 To injectionsneat
                    If NeatSheet.Cells(headerneat + i, CInt(Cell.Column)).Value = "Analyte" Then
                        NeatArea = NeatSheet.Cells(headerneat + i, AreaColumnNeat).Value
                        SpikeArea = SpikeSheet.Cells(headerspike + i, AreaColumnSpike).Value
                        If NeatSheet.Cells(headerneat + i, IDColumn).Value = SpikeSheet.Cells(headerspike + i, IDColumnSpike).Value Then
                            CompSheet.Cells(a, 1).Value = NeatSheet.Cells(headerneat + i, IDColumn).Value
                        Else
                            TransferControlRange = 1
                            Exit Function
                        End If
                        CompSheet.Cells(a, 1).Value = NeatSheet.Cells(headerneat + i, IDColumn).Value
                        CompSheet.Cells(a, 2).Value = NeatArea / (SpikeArea - NeatArea)
                        ConcValue = (CompSheet.Cells(a, 2).Value - CompSheet.Range("L2")) / (CompSheet.Range("K2"))
                        'LLQ kontroll
                        If ConcValue > 25 Then
                            CompSheet.Cells(a, 3).Value = ConcValue
                        Else
                            CompSheet.Cells(a, 3).Value = "NEG"
                        End If
                        If CompSheet.Cells(a, 3).Value = "NEG" Then
                            CompSheet.Cells(a, 4).Value = "PASS"
                            CompSheet.Cells(a, 4).Font.Bold = True
                            CompSheet.Cells(a, 4).Font.Color = RGB(Red:=2, Green:=147, Blue:=70)
                        Else
                            If NeatSheet.Cells(headerneat + i, RFColumn).Value = "NO" And SpikeSheet.Cells(headerspike + i, RFColumnSpike).Value = "NO" Then
                                CompSheet.Cells(a, 4).Value = "PASS"
                                CompSheet.Cells(a, 4).Font.Bold = True
                                CompSheet.Cells(a, 4).Font.Color = RGB(Red:=2, Green:=147, Blue:=70)
                            Else
                                CompSheet.Cells(a, 4).Value = "FAILED"
                                CompSheet.Cells(a, 4).Font.Bold = True
                                CompSheet.Cells(a, 4).Font.Color = RGB(Red:=195, Green:=0, Blue:=4)
                            End If
                        End If
                        a = a + 1
                    End If
                Next
            End If
        Next
        
        CompSheet.Range("A1").EntireColumn.AutoFit
        CompSheet.Range("D1").EntireColumn.AutoFit
        ColIndex = 11
        With CompSheet.Range("A" & CStr(LastContCell + 4) & ":D" & CStr(a - 1))
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlCenter
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = ColIndex
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = ColIndex
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = ColIndex
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = ColIndex
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = ColIndex
            End With
        End With
        
        With CompSheet.Range("A2:O" & CStr(LastContCell))
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlCenter
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = ColIndex
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = ColIndex
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = ColIndex
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = ColIndex
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = ColIndex
            End With
        End With
    Next
    TransferControlRange = controlMap
End Function
Private Sub CalExecutor(ByVal compInt As Integer, ByVal compName As String, ByVal compSheetPos As Integer, ByVal LastContCell As Integer)
    Dim MetaDataSh As Worksheet, CompSheet As Worksheet, Chrt As Chart, yL As Range, xL As Range, LastCal As String
    LastCal = CStr(compInt + compSheetPos)
    Set CompSheet = ThisWorkbook.Sheets(compName)
    Set MetaDataSh = ThisWorkbook.Sheets("MetaData")
    Set Chrt = CompSheet.Shapes.AddChart2.Chart
    Set yL = CompSheet.Range("$B$" & CStr(compSheetPos + 1) & ":$B$" & LastCal)
    Set xL = MetaDataSh.Range("$F$" & CStr(compSheetPos) & ":$F$" & CStr(CInt(LastCal) - 1))
    SlopeValue = Application.WorksheetFunction.Slope(yL, xL)
    InterceptValue = Application.WorksheetFunction.Intercept(yL, xL)
    With Chrt
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        .ChartType = xlXYScatter
        .SeriesCollection.NewSeries
        With .Parent
            .Left = CompSheet.Range("K2").Left
            .Top = CompSheet.Range("K2").Top
            .Height = CompSheet.Range("K2:K" & CStr(LastContCell)).Height
            .Width = CompSheet.Range("K2:O2").Width
        End With
        With .SeriesCollection(1)
            .Name = "Calibrationcurve: " & compName
            .Values = "='" & compName & "'!$B$" & CStr(compSheetPos + 1) & ":$B$" & LastCal
            .XValues = "='MetaData'!$F$" & CStr(compSheetPos) & ":$F$" & CStr(CInt(LastCal) - 1)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
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
    CompSheet.Range("K2").Value = SlopeValue
    CompSheet.Range("L2").Value = InterceptValue
    
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
Dim FirstSheet As Worksheet
Set FirstSheet = ThisWorkbook.Sheets("Utvärderingsprogram uTAC-Bens")
'Synkroniserar Sheet1s namn med cell Target
    If Target.Address <> "$A$1" Then Exit Sub
    If IsEmpty(Target) Then Exit Sub
    If Len(Target.Value) > 31 Then
        Application.EnableEvents = False
        Target.ClearContents
        Application.EnableEvents = True
        Exit Sub
    End If
    FirstSheet.Cells(1, 1).Value = "Utvärderingsprogram uTAC-Bens"
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

    Dim strSheetName As String, CtrlWSh As Worksheet, bln As Boolean
    strSheetName = Trim(Target.Value)
    On Error Resume Next
    Set CtrlWSh = ActiveWorkbook.Worksheets(strSheetName)
    On Error Resume Next
    If Not CtrlWSh Is Nothing Then
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
