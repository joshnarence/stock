Public Sub stock()

' Delete columns J:R containing old results

For Each ws In Worksheets
ws.Columns("J:R").EntireColumn.Delete
Next ws

' Looping through every worksheet to process data

For Each ws In Worksheets
Dim worksheetname As String

' Finding the last rows of data in the worksheet

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Insert columns J:R to populate results

ws.Range("J1").EntireColumn.Insert
ws.Cells(1, 10).Value = "Ticker"
ws.Range("K1").EntireColumn.Insert
ws.Cells(1, 11).Value = "yearly change"
ws.Range("L1").EntireColumn.Insert
ws.Cells(1, 12).Value = "percent change"
ws.Range("M1").EntireColumn.Insert
ws.Cells(1, 13).Value = "Total volume"
ws.Range("P1").EntireColumn.Insert
ws.Cells(1, 16).Value = ""
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
ws.Range("Q1").EntireColumn.Insert
ws.Cells(1, 17).Value = "Ticker"
ws.Range("R1").EntireColumn.Insert
ws.Cells(1, 18).Value = "Value"


'Assign variables

Dim j As Double
Dim openprice As Double
Dim closeprice As Double
Dim percent As Double
Dim totalvolume As Double
Dim xmax As Double
Dim xmin As Double
Dim xmaxv As Double
Dim xmaxtick As String
Dim xmintick As String
Dim xmaxvtick As String


totalvolume = 0
openprice = ws.Cells(2, 3).Value

j = 2

'Looping through all the rows of the worksheet

For i = 2 To lastrow

'Adding the volume of unique Ticker

totalvolume = ws.Cells(i, 7) + totalvolume
    
'check if the next row has a different Ticker
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(j, 10).Value = ws.Cells(i, 1).Value
        closeprice = ws.Cells(i, 6).Value
        ws.Cells(j, 11).Value = closeprice - openprice
   
'Assign colors based on yearly change results

            If ws.Cells(j, 11).Value < 0 Then
                 ws.Cells(j, 11).Interior.ColorIndex = 3
            Else
                 ws.Cells(j, 11).Interior.ColorIndex = 43
            End If
        
        
        percent = ws.Cells(j, 11).Value / openprice

        ws.Cells(j, 12).Value = percent
        openprice = ws.Cells(i + 1, 3).Value
        
'Insert value of total volume computed for unique Ticker

        ws.Cells(j, 13).Value = totalvolume
        

        totalvolume = 0
        j = j + 1
    End If
   

Next i

' Find the last row of the unique Ticker Column
lastrowk = ws.Cells(Rows.Count, 10).End(xlUp).Row
xmax = ws.Cells(2, 12).Value
xmaxtick = ws.Cells(2, 1).Value
xmin = ws.Cells(2, 12).Value
xmintick = ws.Cells(2, 1).Value
xmaxv = ws.Cells(2, 13).Value
xmaxvtick = ws.Cells(2, 1).Value

' Looping through unique Ticker Rows to calculate Greatest % increase, % decrease and Greatest Total Volume

For k = 2 To lastrowk

    If ws.Cells(k, 12).Value > xmax Then
    xmax = ws.Cells(k, 12).Value
    xmaxtick = ws.Cells(k, 1).Value
    End If

    If ws.Cells(k, 12).Value < xmin Then
    xmin = ws.Cells(k, 12).Value
    xmintick = ws.Cells(k, 1).Value
    End If

    If ws.Cells(k, 13).Value > xmaxv Then
    xmaxv = ws.Cells(k, 13).Value
    xmaxvtick = ws.Cells(k, 1).Value
    End If

Next k

ws.Cells(2, 18).Value = xmax
ws.Cells(3, 18).Value = xmin
ws.Cells(4, 18).Value = xmaxv

ws.Cells(2, 17).Value = xmaxtick
ws.Cells(3, 17).Value = xmintick
ws.Cells(4, 17).Value = xmaxvtick

' Convert percent change and greatest % changes to percentage format

ws.Range("L2:L" & lastrowk).NumberFormat = "0.00%"
ws.Cells(2, 18).NumberFormat = "0.00%"
ws.Cells(3, 18).NumberFormat = "0.00%"


Next ws
End Sub

