Attribute VB_Name = "ARG"
Option Explicit

Sub pred_returns()

    Dim pred As Range
    Dim n As Integer
    Dim êta As Range
    Dim price As Range
    
   Dim g As Worksheet
   Set g = ThisWorkbook.Sheets("Feuil6")
   
   Dim i As Integer
   Dim k As Integer
   Dim sigma As Range
   Dim epsil As Range
   
   Dim sum As Integer
   epsil = g.Range("M2")
   sigma = g.Range("L2")
   n = g.Cells(2, 10)
   êta = g.Range("K2")
   price = g.Range("A2")
   For i = 1 To n
   
        For k = 1 To n
        
            sum = sum + (êta(k) * price(i - 1 - k))
            
        Next k
        
            pred(i) = sum + Sqr(sigma(i) * epsil(i))
        
    Next i
    
 g.Range("C2") = pred
 

End Sub
