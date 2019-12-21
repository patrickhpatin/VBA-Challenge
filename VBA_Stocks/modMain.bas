Attribute VB_Name = "modMain"
Private Type typSummaryRow
    Ticker      As String
    OpenRowNum  As Long
    OpenValue   As Long
    CloseRowNum As Long
    CloseValue  As Long
End Type

Private oTicker                     As typSummaryRow

Private m_lGreatestIncreaseRow      As Long
Private m_lGreatestDecreaseRow      As Long
Private m_lGreatestTotalVolumeRow   As Long


' Keeps running totals for each ticker symbol and
' writes to the cells I - L and N - P
Public Sub AggregateWorksheets()
    On Error GoTo ErrHand
    
    Dim mySheet As Worksheet
    
    For Each mySheet In ActiveWorkbook.Sheets
        ' This next line is for testing only
'        Set mySheet = ActiveWorkbook.Sheets("2016")
        mySheet.Select
        AddHeaderOnWorksheet
        PopulateSummaryTable
        PopulateHighlightTable
    Next
    
    MsgBox "The aggregation has completed successfully!"
    
Exit_Clean:
    On Error Resume Next
    
    Set mySheet = Nothing
    
    Exit Sub
ErrHand:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error: " & Err.Number, Err.HelpFile, Err.HelpContext
    Resume Exit_Clean
End Sub


Private Sub PopulateSummaryTable()
    On Error GoTo ErrHand
    
    Dim lRow            As Long
    Dim lRowCount       As Long
    Dim iSummaryRow     As Integer
    Dim mySheet         As New Worksheet
    Dim sCurrentTicker  As String
    
    ' Set the active sheet and initialize variables
    iSummaryRow = 1
    Set mySheet = ActiveWorkbook.ActiveSheet
    ' I want to loop to the first empty record
    ' that way the last ticker value is also recorded
    lRowCount = mySheet.UsedRange.Rows.Count + 1
    
    With mySheet
        For lRow = 2 To lRowCount
            ' Get information from the current row.
            sCurrentTicker = .Range("A" & lRow).Value
            
            ' Was this row and the last row the same stock?
            If lRow = 2 Then
                ' Reset these variables for every sheet
                ' These represent Summary Rows
                m_lGreatestIncreaseRow = lRow
                m_lGreatestTotalVolumeRow = lRow
                m_lGreatestDecreaseRow = lRow
                
                oTicker.Ticker = .Range("A" & lRow).Value
                oTicker.OpenRowNum = lRow
                oTicker.OpenValue = .Range("C" & lRow).Value
                oTicker.CloseValue = 0
                oTicker.CloseRowNum = 2
            ElseIf (StrComp(sCurrentTicker, oTicker.Ticker, vbTextCompare) = 0) Then ' Same
                oTicker.CloseRowNum = lRow
                oTicker.CloseValue = .Range("F" & lRow).Value
            Else ' Different
                iSummaryRow = iSummaryRow + 1
                ' Ticker Symbol
                .Range("I" & iSummaryRow).Value = oTicker.Ticker
                
                ' Yearly Change
                .Range("J" & iSummaryRow).Formula = "=F" & oTicker.CloseRowNum & " - C" & oTicker.OpenRowNum
                
                ' Is the percent change negative?
                ' If so, make that value Red
                ' Otherwise, make it green
                .Range("J" & iSummaryRow).Select
                If .Range("J" & iSummaryRow).Value > 0 Then
                    .Range("J" & iSummaryRow).Interior.ColorIndex = 4 ' Green
                Else
                    .Range("J" & iSummaryRow).Interior.ColorIndex = 3 ' Red
                End If
                
                ' Percent Change
                If .Range("C" & oTicker.OpenRowNum).Value = 0 Then
                    .Range("K" & iSummaryRow).Formula = 0
                Else
                    .Range("K" & iSummaryRow).Formula = "=J" & iSummaryRow & " / C" & oTicker.OpenRowNum
                End If
                
                ' Total Volume
                .Range("L" & iSummaryRow).Formula = "=SUM(G" & oTicker.OpenRowNum & ":G" & oTicker.CloseRowNum & ")"
                
                ' Check for Highlight Records
                ' Greatest Percent Increase
                If .Range("K" & iSummaryRow).Value > .Range("K" & m_lGreatestIncreaseRow).Value Then
                    m_lGreatestIncreaseRow = iSummaryRow
                End If
                
                ' Greatest Percent Decrease
                If .Range("K" & iSummaryRow).Value < .Range("K" & m_lGreatestDecreaseRow).Value Then
                    m_lGreatestDecreaseRow = iSummaryRow
                End If
                
                ' Greatest Total Volume
                If .Range("L" & iSummaryRow).Value > .Range("L" & m_lGreatestTotalVolumeRow).Value Then
                    m_lGreatestTotalVolumeRow = iSummaryRow
                End If
                
                ' Reset Ticker Structure
                oTicker.Ticker = .Range("A" & lRow).Value
                oTicker.OpenRowNum = lRow
                oTicker.OpenValue = .Range("C" & lRow).Value
                oTicker.CloseRowNum = 0
                oTicker.CloseValue = 0
            End If
        Next lRow
        
        ' Format Columns
        Range("J2:J" & iSummaryRow).Select
        Selection.Style = "Currency"
        
        Range("K2:K" & iSummaryRow).Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
        
        Range("L2:L" & iSummaryRow).Select
        Selection.Style = "Comma"
        Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    End With

Exit_Clean:
    On Error Resume Next
    
    Set mySheet = Nothing
    
    Exit Sub
ErrHand:
    Err.Raise Err.Number, "Populate Summary Table", Err.Description, Err.HelpFile, Err.HelpContext
    GoTo Exit_Clean
End Sub


Private Sub PopulateHighlightTable()
    With ActiveWorkbook.ActiveSheet
        ' Set Ticker Symbols
        .Range("O2").Value = .Range("I" & m_lGreatestIncreaseRow).Value
        .Range("O3").Value = .Range("I" & m_lGreatestDecreaseRow).Value
        .Range("O4").Value = .Range("I" & m_lGreatestTotalVolumeRow).Value
        
        ' Greatest Percent Increase
        .Range("P2").Value = .Range("K" & m_lGreatestIncreaseRow).Value
        .Range("P2").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
        
        ' Greatest Percent Decrease
        .Range("P3").Value = .Range("K" & m_lGreatestDecreaseRow).Value
        .Range("P3").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
        
        ' Greatest Total Volume
        .Range("P4").Value = .Range("L" & m_lGreatestTotalVolumeRow).Value
        .Range("P4").Select
        Selection.Style = "Comma"
        Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        
        ' Autofit the columns so we can see all of the data
        Columns("I:P").EntireColumn.AutoFit
    End With
End Sub

Sub test()
    Dim lNum As Long
    lNum = Range("P4").Value
    
    MsgBox CStr(lNum)
End Sub


Private Sub AddHeaderOnWorksheet()
    Dim lRowCount As Long
    
    With ActiveWorkbook.ActiveSheet
        ' First make sure the data is sorted by:
        '      a.) Ticker (Alphabetical)
        '      b.) Date (Ascending)
        Cells.Select
        .Sort.SortFields.Clear
        lRowCount = .UsedRange.Rows.Count
        .Sort.SortFields.Add Key:=Range("A2:A" & lRowCount), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("B2:B" & lRowCount), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range("A1:G" & lRowCount)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ' Headers for Aggregation Rows
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "% Change"
        .Range("L1").Value = "Total Volume"
        
        ' Headers for Highlight Rows
        .Range("O1").Value = "Ticker"
        .Range("P1").Value = "Value"
        .Range("N2").Value = "Greatest % Increase:"
        .Range("N3").Value = "Greatest % Decrease:"
        .Range("N4").Value = "Greatest Total Volume:"
        
        .Range("I1:L1, N1:P1, N1:N4").Select
        Selection.Font.Bold = True
        .Range("I1:L1, N1:P1").Select
        Selection.HorizontalAlignment = xlCenter
    End With
End Sub














