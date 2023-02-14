'
'CD 2022-08-27
'Excel VBA to cycle each product through the forecast model, and post results in separate sheet
'
Sub runForecast()
    Dim StartTime As Double
    Dim SecondsElapsed As String
	'start timer for runtime
    StartTime = Timer
    Dim c As Range
    Dim destRange As Range
	'set starting point for result arrays
    Set destRange = ThisWorkbook.Sheets("Results").Range("A2")
    Dim skuRange As Range
	'set target cell for all the formulas
    Set skuRange = ThisWorkbook.Sheets("model_ex").Range("A1")
    Dim weekRange As Range
	'set starting range for forcast results per product
    Set weekRange = ThisWorkbook.Sheets("model_ex").Range("J104:J129")
	'This range contains the distinct skus and may need resized
    For Each c In ThisWorkbook.Sheets("Sku").Range("A1:A1000")
        'exit routine for last sku	
        If c.Value2 = "" Then
            Exit For
        End If
		'disregards header if exists, due to user error
        If c.Value2 <> "product_code" Then
            skuRange.Value2 = c.Value2
            For Each wk In weekRange
			    'sku array
                destRange.Value2 = c.Value2
				'isoWeekNum array
                destRange.offset(0, 1).Value2 = wk.Value2
				'forecast qty array
                destRange.offset(0, 2).Value2 = wk.offset(0, 3).Value2
				'weekPeriod array
                destRange.offset(0, 3).Value2 = wk.offset(0, 6).Value2
				'past 26w actual sales array
                destRange.offset(0, 4).Value2 = wk.offset(-26, 1).Value2
				'reset starting point for new results
                Set destRange = destRange.offset(1, 0)
            Next wk
        End If
    Next c
	'convert total runtime to minutes
    SecondsElapsed = Round((Timer - StartTime) / 60, 2)
    MsgBox "Forecast Completed, See Results Tab. Run Time " & SecondsElapsed & " minutes"
End Sub