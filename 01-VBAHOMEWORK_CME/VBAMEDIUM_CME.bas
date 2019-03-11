Attribute VB_Name = "Module2"
Sub VBAmedium()
'Medium part of VBA homework assignment

' Definitions  and Declarations
Dim i As Double  ' Row being read by loop at current
Dim j As Double  ' updated value of i modified by loop at current
Dim k As String  ' recorded value of current <ticker>
Dim n As Double  ' records value to use as starting point for outputs
Dim total As Double  ' recorded value of current <ticker>
Dim start_value As Double  ' recorded starting value of newest <ticker>
Dim next1 As Double  ' used to reset when finding new start_value
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

' Headers for columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

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
           
            i = i + 1  ' advance read to next Row
            j = i
            n = n + 1  ' advance output to next Row
            total = 0  ' reset total to zero
            next1 = 0  ' reset next1 to zero to get new start_value

        End If
        
        End If

    Next iLoop

End Sub

