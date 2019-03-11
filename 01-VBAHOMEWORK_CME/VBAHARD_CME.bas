Attribute VB_Name = "Module3"
Sub VBAhard()
'Hard part of VBA homework assignment

' Definitions  and Declarations
Dim i As Double  ' Row being read by loop at current
Dim j As Double  ' updated value of i modified by loop at current
Dim k As String  ' recorded value of current <ticker>
Dim n As Double  ' records value to use as starting point for outputs
Dim total As Double  ' recorded value of current <ticker>
Dim start_value As Double  ' recorded starting value of newest <ticker>
Dim next1 As Double  ' used to reset when finding new start_value
Dim max As Double
Dim min As Double
Dim max_total As Double ' Updated highest value of total at current
Dim lastrow As Long  ' Last Row in spreadsheet

lastrow = Range("A1").End(xlDown).Row ' Defined range of lastrow

' all start at 2 (modified later by current step in loop)
i = 2
j = 2
k = 2
n = 2

' all start at 0 (modified later by current step in loop)
total = 0
next1 = 0
max = 0
min = 0
max_total = 0

' Headers for columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"


' LOOP THROUGH EACH WORKSHEET
' For Each ws In Worksheets
   ' ws.Activate
   
    ' Primary loop, runs through whole sheet.
    For iLoop = 2 To lastrow
    
        ' Start gathering values based on <ticker>
        ' Determine current <ticker> value and calculate total rolling update
        If Cells(i, 1).Value = Cells(j + 1, 1).Value Then
            ' Determine <open> value of current <ticker>
            If next1 = 0 Then
                start_value = Cells(i, 3).Value
                next1 = 1
            End If

          k = Cells(j, 1).Value
          total = total + Cells(i, 7).Value
          i = i + 1  ' advance read to next Row
          j = i

        ' Next cells changes, calculate final total of current <ticker> and output
        Else
          Cells(n, 9).Value = Cells(i, 1).Value    ' Outputting <ticker> for record of "Ticker"
          total = total + Cells(i, 7).Value    ' Calculating final total value for current <ticker>
          Cells(n, 12).Value = total    ' Outputting value of total to "Total Stock Volume"

            ' Find "Greatest Total Volume" and output
            If total > max_total Then  ' Compare against current max
               max_total = total  ' Record newest high value
               Cells(4, 15).Value = Cells(j, 1).Value
               Cells(4, 16).Value = max_total
               Cells(4, 14).Value = "Greatest Total Volume"

            End If
        
                ' Calculating values for range for "Year Change"
                Cells(n, 10).Value = Cells(i, 6).Value - start_value
                '
                If Cells(n, 10).Value > "0" Then
                   Cells(n, 10).Interior.ColorIndex = 4

                Else
                   Cells(n, 10).Interior.ColorIndex = 3

                End If
            
                'cell formatting
                If Cells(n, 10).Value = 0 Then
                   Cells(n, 11).Value = 0
                   Cells(n, 11).NumberFormat = "0.00%"

                End If
            
                ' rectifying values for formatting for "Percent Change"
                If CDbl(start_value) = 0 Or Cells(n, 10).Value = 0 Then
                   Cells(n, 11).Value = 0
                   Cells(n, 11).NumberFormat = "0.00%"

                End If
        
        ' Determine min and max and format cells
        If start_value <> 0 Then
           Cells(n, 11).Value = Cells(n, 10).Value / start_value
           Cells(n, 11).NumberFormat = "0.00%"
            ' max
            If Cells(n, 11).Value > max Then  ' Determine highest % difference in <close> / <open>
               max = Cells(n, 11).Value  ' Update max to new highest value
               Cells(2, 16).Value = max
               Cells(2, 16).NumberFormat = "0.00%"
               Cells(2, 15).Value = Cells(j, 1).Value
               Cells(2, 14).Value = "Greatest % Increase"

            End If
            ' min
            If Cells(n, 11).Value < min Then  ' Determine lowest % difference in <close> / <open>
               min = Cells(n, 11).Value  ' Update min to new lowest value
              Cells(3, 16).Value = min  ' Output followed by formatting and caption
              Cells(3, 16).NumberFormat = "0.00%"
              Cells(3, 15).Value = Cells(j, 1).Value
              Cells(3, 14).Value = "Greatest % Decrease"
        
            End If

        End If

            i = i + 1  ' advance read to next Row
            j = i
            n = n + 1  ' advance output to next Row
            total = 0  ' reset total to zero
            next1 = 0  ' reset next1 to zero to get new start_value

        End If  ' This formatting got a bit sloppy, sorry about that.

    Next iLoop

End Sub


