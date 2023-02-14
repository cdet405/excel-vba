'
'CD 2023-02-06 version 2 
'Excel VBA to cycle each product through the forecast model, and post results in separate sheet
'
Sub runForecast()
    Dim StartTime As Double
    Dim SecondsElapsed As String
	'start timer for runtime
    StartTime = Timer

    'set target cell for sku in of all of the forecast calcs
	Dim skuRange As Range
    Set skuRange = ThisWorkbook.Sheets("model_ex").Range("A1")

    'Range of distinct Skus to cycle through - resize if needed
    Dim skuList As Range
    Set skuList = ThisWorkbook.Sheets("Sku").Range("A1:A1000")

    'starting point for where to post result array
    Dim resultsList As Range
    Set resultsList = ThisWorkbook.Sheets("Results").Range("A2")

    'starting point for forecast result array
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
    ReDim resultsData(1 To skuRows * weekRows, 1 To 5)

    j = 1
    For i = 1 To skuRows
	    'exit routine when all skus have been cycled through
        If skuData(i, 1) = "" Then
            Exit For
        End If
		'disregard header if exists
        If skuData(i, 1) <> "product_code" Then
            skuRange.Value2 = skuData(i, 1)
            wkMatch = Application.WorksheetFunction.Match(skuData(i, 1), skuList, 0)
            For k = 1 To weekRows
			    'sku
                resultsData(j, 1) = skuData(i, 1)
				'isoWeekNum
                resultsData(j, 2) = weekRange.Value2(k, 1)
				'qty
                resultsData(j, 3) = weekRange.Cells(k, 4).Value2
				'weekPeriod
                resultsData(j, 4) = weekRange.Cells(k, 7).Value2
                'past 26w actual sales
                resultsData(j, 5) = weekRange.Cells(k - 26, 2).Value2
                j = j + 1
            Next k
        End If
    Next i

    'Write the results to the worksheet
    resultsList.Resize(skuRows * weekRows, 5).Value2 = resultsData

    SecondsElapsed = Round((Timer - StartTime) / 60, 2)
    MsgBox "Forecast Completed, See Results Tab. Run Time " & SecondsElapsed & " minutes"

    'Turn on screen updating
    Application.ScreenUpdating = True
End Sub

