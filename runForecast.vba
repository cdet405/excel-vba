' +-----------------------------------------+
' | version 3 - CD 2023-04-24               |
' | for use in forecast & vf_forecast .xlsm |
' +-----------------------------------------+

Sub runForecast()
    Dim StartTime As Double
    Dim SecondsElapsed As String
    StartTime = Timer

    Dim skuRange As Range
    Set skuRange = ThisWorkbook.Sheets("model_ex").Range("A1")

    Dim skuList As Range
    Set skuList = ThisWorkbook.Sheets("Sku").Range("A1:A7500")

    Dim resultsList As Range
    Set resultsList = ThisWorkbook.Sheets("Results").Range("A2")

    Dim weekRange As Range
    Set weekRange = ThisWorkbook.Sheets("model_ex").Range("J104:J129")

    Dim wkMatch As Integer
    Dim skuData As Variant
    Dim resultsData As Variant
    Dim i As Long, j As Long

    'Turn off screen updating to improve performance
    Application.ScreenUpdating = False

    'Get the list of SKUs
    skuData = skuList.Value2

    'Get the number of rows in the SKU list
    Dim skuRows As Long
    skuRows = UBound(skuData, 1)

    'Get the number of weeks in the week list
    Dim weekRows As Long
    weekRows = UBound(weekRange.Value2, 1)

    'Allocate memory for the results data
    ReDim resultsData(1 To skuRows * weekRows, 1 To 9)

    j = 1
    For i = 1 To skuRows
        If skuData(i, 1) = "" Then
            Exit For
        End If
        If skuData(i, 1) <> "product_code" Then
            skuRange.Value2 = skuData(i, 1)
            wkMatch = Application.WorksheetFunction.Match(skuData(i, 1), skuList, 0)
            For k = 1 To weekRows
                resultsData(j, 1) = skuData(i, 1) 'sku
                resultsData(j, 2) = weekRange.Value2(k, 1) 'week
                resultsData(j, 3) = weekRange.Cells(k, 4).Value2 'qty
                resultsData(j, 4) = weekRange.Cells(k, 7).Value2 'weekPeriod
                resultsData(j, 5) = weekRange.Cells(k - 26, 2).Value2 'LTM
                resultsData(j, 6) = weekRange.Cells(k - 26, 1).Value2 'LTMweek
                resultsData(j, 7) = weekRange.Cells(k - 26, 11).Value2 'LTMoutlier
                resultsData(j, 8) = weekRange.Cells(k - 26, 12).Value2 'LTMsmooth
                resultsData(j, 9) = weekRange.Cells(k, 8).Value2 'Szn Index
                'Debug.Print "Processing Sku Batch For: " & resultsData(j, 1)
                j = j + 1
            Next k
        End If
    Next i

    'Write the results to the worksheet
    resultsList.Resize(skuRows * weekRows, 9).Value2 = resultsData

    SecondsElapsed = Round((Timer - StartTime) / 60, 2)
    MsgBox "Forecast Completed, See Results Tab. Run Time " & SecondsElapsed & " minutes"

    'Turn on screen updating
    Application.ScreenUpdating = True
End Sub

