' Runs Each Sku through formula to create a master sheet 
Sub generateBase()
Dim StartTime As Double
Dim SecondsElapsed As String
StartTime = Timer

Dim SkuRange As Range
Set SkuRange = ThisWorkbook.Sheets("main").Range("A1")

Dim SkuList As Range
Set SkuList = ThisWorkbook.Sheets("sku").Range("A1:A1000")

    Dim resultsList As Range
    Set resultsList = ThisWorkbook.Sheets("result").Range("A2")

    Dim keyRange As Range
    Set keyRange = ThisWorkbook.Sheets("main").Range("A3:A43")

    Dim keyMatch As Integer
    Dim skuData As Variant
    Dim resultsData As Variant
    Dim i As Long, j As Long

    'Turn off screen updating to improve performance
    Application.ScreenUpdating = False

    'Get the list of SKUs
    skuData = SkuList.Value2

    'Get the number of rows in the SKU list
    Dim skuRows As Long
    skuRows = UBound(skuData, 1)

    'Get the number of rows in the key list
    Dim keyRows As Long
    keyRows = UBound(keyRange.Value2, 1)

    'Allocate memory for the results data
    ReDim resultsData(1 To skuRows * keyRows, 1 To 10)

    j = 1
    For i = 1 To skuRows
        If skuData(i, 1) = "" Then
            Exit For
        End If
        If skuData(i, 1) <> "product_code" Then
            SkuRange.Value2 = skuData(i, 1)
            keyMatch = Application.WorksheetFunction.Match(skuData(i, 1), SkuList, 0)
            For k = 1 To keyRows
                resultsData(j, 1) = skuData(i, 1) 'sku
                resultsData(j, 2) = keyRange.Value2(k, 1) 'time
                resultsData(j, 3) = keyRange.Cells(k, 2).Value2 'cci
                resultsData(j, 4) = keyRange.Cells(k, 3).Value2 'cpi
                resultsData(j, 5) = keyRange.Cells(k, 4).Value2 'cpimm
                resultsData(j, 6) = keyRange.Cells(k, 5).Value2 'ffcpi
                resultsData(j, 7) = keyRange.Cells(k, 6).Value2 'sales
                resultsData(j, 8) = keyRange.Cells(k, 7).Value2 'date
                resultsData(j, 9) = keyRange.Cells(k, 8).Value2 'sid
                resultsData(j, 10) = keyRange.Cells(k, 9).Value2 'pid
                j = j + 1
            Next k
        End If
    Next i

    'Write the results to the worksheet
    resultsList.Resize(skuRows * keyRows, 10).Value2 = resultsData

    SecondsElapsed = Round((Timer - StartTime), 2)
    MsgBox "Script Completed, See Result Tab. Run Time " & SecondsElapsed & " Seconds"

    'Turn on screen updating
    Application.ScreenUpdating = True
End Sub
