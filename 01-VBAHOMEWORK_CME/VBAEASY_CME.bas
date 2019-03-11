Attribute VB_Name = "Module1"
Sub VBAeasy()
'Easy part of VBA homework assignment

' Definitions  and Declarations
' Dim ws As Worksheet ' To work through Worksheets
Dim i As Double  ' Row being read by loop at current
Dim j As Double  ' updated value of i modified by loop at current
Dim k As String  ' recorded value of current <ticker>
Dim n As Double  ' records value to use as starting point for outputs
Dim total As Double  ' recorded value of current <ticker>
Dim lastrow As Long  ' Last Row in spreadsheet

lastrow = Range("A1").End(xlDown).Row ' Defined range of lastrow

' all start at 2 (modified later by current step in loop)
i = 2
j = 2
k = 2
n = 2

' all start at 0 (modified later by current step in loop)
total = 0

' Headers for columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"

' LOOP THROUGH EACH WORKSHEET
' For Each ws In Worksheets
   ' ws.Activate
    
    ' Primary loop, runs through whole sheet.
    For iLoop = 2 To lastrow
    
        ' Start gathering values based on <ticker>
        ' Determine current <ticker> value and calculate total rolling update
        If Cells(i, 1).Value = Cells(j + 1, 1).Value Then
          k = Cells(j, 1).Value
          total = total + Cells(i, 7).Value
          i = i + 1  ' advance read to next Row
          j = i
        
        ' Next cells changes, calculate final total of current <ticker> and output
        Else
          Cells(n, 9).Value = Cells(i, 1).Value    ' Outputting <ticker> for record of "Ticker"
          total = total + Cells(i, 7).Value    ' Calculating final total value for current <ticker>
          Cells(n, 10).Value = total    ' Outputting value of total to "Total Stock Volume"
          
          i = i + 1  ' advance read to next Row
          j = i
          n = n + 1  ' advance output to next Row
          total = 0  ' reset total to zero

        End If
    
    Next iLoop

' Next ws

End Sub


