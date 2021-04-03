Attribute VB_Name = "Preparation_of_data"
Option Explicit
Sub main()
    Returns
    features
    SMA
    standard
    Upper_Lower

End Sub


Sub Returns() ' le sub qui donne les returns

    
    Dim NumRows As Integer
    Dim f As Worksheet
    Dim i As Integer
    Set f = ThisWorkbook.Sheets("Feuil1")
   
    NumRows = f.Range("A2", Range("A2").End(xlDown)).Rows.Count
    Debug.Print (NumRows)
    
    Dim vartest As Variant
    
    For i = 1 To NumRows - 1
    
        f.Cells(i + 1, 2).Value = (f.Cells(i + 2, 1).Value - f.Cells(i + 1, 1).Value) / f.Cells(i + 1, 1).Value
    
    Next i
    

End Sub

Sub features() ' le sub qui donne les Y
    
    Dim NumRows As Integer
    Dim f As Worksheet
    Dim i As Integer
    Set f = ThisWorkbook.Sheets("Feuil1")
   
    NumRows = f.Range("B2", Range("B2").End(xlDown)).Rows.Count
    Debug.Print (NumRows)
    
    For i = 2 To NumRows - 1
    
         If (f.Cells(i, 2) < 0) Then
         
            f.Cells(i, 3) = 0
         ElseIf (f.Cells(i, 2) > 0) Then
            
            f.Cells(i, 3) = 1
            
         End If
         
         
    
    Next i
    
End Sub

Sub SMA() ' sub donnant la moyenne mobile

    Dim NumRows As Integer
    Dim f As Worksheet
    
    Set f = ThisWorkbook.Sheets("Feuil1")
    Dim window As Integer

    window = 20
    NumRows = f.Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    Dim a As Double
    Dim k As Integer
    Dim i As Integer
    
   
   
   
   For k = 0 To 1741
    
        Dim som As Double
     
        som = moy(k) / 20
        
     f.Cells(21 + k, 4) = som
    
  Next k
   
        
      
End Sub


Function moy(a As Integer) ' allow us to compute the SMA sub

    Dim i As Integer
    Dim som As Double
     Dim NumRows As Integer
    Dim f As Worksheet
    
    Set f = ThisWorkbook.Sheets("Feuil1")
    Dim window As Integer

    window = 20
    NumRows = f.Range("A2", Range("A2").End(xlDown)).Rows.Count
        
        For i = 2 + a To 21 + a
        
            som = som + (f.Cells(i, 2))
    
        Next i
   
    moy = som


End Function


Function stand(a As Integer) ' allow us to compute the standard deviation
    
    Dim i As Integer
    Dim som As Double
     Dim NumRows As Integer
    Dim f As Worksheet
    
    Set f = ThisWorkbook.Sheets("Feuil1")
    Dim window As Integer

    window = 20
    NumRows = f.Range("A2", Range("A2").End(xlDown)).Rows.Count
        
        For i = 2 + a To 21 + a
        
            som = som + (f.Cells(i, 2) - f.Cells(a + 21, 4)) * (f.Cells(i, 2) - f.Cells(a + 21, 4))
    
        Next i
   
    stand = som
    
End Function


Sub standard() ' allow to fill the tab of standard deviation

    
    Dim i As Integer
    
     Dim NumRows As Integer
    Dim f As Worksheet
    
    Set f = ThisWorkbook.Sheets("Feuil1")
    Dim window As Integer

    window = 20
    NumRows = f.Range("A2", Range("A2").End(xlDown)).Rows.Count
    Dim k As Integer
    
    For k = 0 To 1741
    
        Dim som As Double
     
        som = stand(k) / 20
        
     f.Cells(21 + k, 5) = som
    
  Next k

End Sub



Sub Upper_Lower()

    Dim i As Integer
    
    Dim NumRows As Integer
    Dim f As Worksheet
    
    Set f = ThisWorkbook.Sheets("Feuil1")
    Dim window As Integer

    window = 20
    NumRows = f.Range("A2", Range("A2").End(xlDown)).Rows.Count
    Dim k As Integer
    
    
    For i = 21 To NumRows
    
        f.Cells(i, 6) = f.Cells(i, 4) + 2 * f.Cells(i, 5)
        f.Cells(i, 7) = f.Cells(i, 4) - 2 * f.Cells(i, 5)
    
    Next i

End Sub



