Attribute VB_Name = "Multiple_perceptron"
Option Explicit



Sub main3()
    MLinear_P
    
    
    
End Sub

Sub MLinear_P() 'ici n c'est le nombre de perceptron
    ' first Perceptron
    Dim P1 As New Perceptron
    Dim LR As Double
    Dim NI As Integer
    P1.set_Rate = LR
    P1.set_Iter = NI
    Dim W1() As Double
    W1 = P1.P_Weights(3)
    
   ' second Perceptron
   
   Dim P2 As New Perceptron
    P2.set_Rate = LR
    P2.set_Iter = NI
    Dim W2() As Double
    W2 = P2.P_Weights(3)
    
    'THIRD perceptron
    
    Dim P3 As New Perceptron
    P3.set_Rate = LR
    P3.set_Iter = NI
    Dim W3() As Double
    W3 = P3.P_Weights(3)
    
   ' fill the excel with weights values
   Dim k As Integer
   Dim j As Integer
   Dim f As Worksheet
    Set f = ThisWorkbook.Sheets("Feuil1")
    Dim g As Worksheet
    Set g = ThisWorkbook.Sheets("Feuil4")
   For k = 1 To 3
        g.Cells(2, k + 4 - 1) = W1(k)
        g.Cells(3, k + 4 - 1) = W2(k)
        g.Cells(4, k + 4 - 1) = W3(k)
        
    Next k
   ' biais initial
   Dim b1 As Double
   Dim b2 As Double
   Dim b3 As Double
   b1 = 0
   b2 = 0
   b3 = 0
   g.Cells(2, 7) = b1
   g.Cells(3, 7) = b2
   g.Cells(4, 7) = b3
   
   ' Preparation of the Xi
    
    Dim X1() As Double
    ReDim X1(1 To 1743)                  'SMA standard deviation
    Dim X2() As Double  'UP
    ReDim X2(1 To 1743)
    Dim X3() As Double  'LB
    ReDim X3(1 To 1743)
    
   For k = 21 To 1763
        
        X1(k - 20) = f.Cells(k, 5)
        X2(k - 20) = f.Cells(k, 6)
        X3(k - 20) = f.Cells(k, 7)
        
    
    Next k
    ' remplissage des cellules de la feuille2
    
    For k = 2 To UBound(X1) + 1
        
        g.Cells(k, 1) = X1(k - 1)
        g.Cells(k, 2) = X2(k - 1)
        g.Cells(k, 3) = X3(k - 1)
        
    
    Next k
   
   
   ' Preparation of the Z
   
    Dim Z1() As Double
    ReDim Z1(1 To 1743)
    Dim Z2() As Double
    ReDim Z2(1 To 1743)
    Dim Z3() As Double
    ReDim Z3(1 To 1743)
    
     For k = 1 To UBound(X1)
    
        Z1(k) = W1(1) * X1(k) + W1(2) * X2(k) + W1(3) * X3(k) + b1
        Z2(k) = W2(1) * X1(k) + W2(2) * X2(k) + W2(3) * X3(k) + b2
        Z3(k) = W3(1) * X1(k) + W3(2) * X2(k) + W3(3) * X3(k) + b3
        g.Cells(k + 1, 8) = Z1(k)
        g.Cells(k + 1, 9) = Z2(k)
        g.Cells(k + 1, 10) = Z3(k)
        
        If (Z1(k) > 0) Then
            g.Cells(k + 1, 11) = 1
        Else
            g.Cells(k + 1, 11) = 0
        End If
        
         If (Z2(k) > 0) Then
            g.Cells(k + 1, 12) = 1
        Else
            g.Cells(k + 1, 12) = 0
        End If
         
         If (Z3(k) > 0) Then
            g.Cells(k + 1, 13) = 1
        Else
            g.Cells(k + 1, 13) = 0
        End If
        
        
    Next k
 
        ' calculate the activate
        
        ' Valeurs prédites avec fonctions d'activations Y^
            
            'Preparation of Ai
            Dim A1() As Double
            ReDim A1(1 To 1743)
            Dim A2() As Double
            ReDim A2(1 To 1743)
            Dim A3() As Double
            ReDim A3(1 To 1743)
   
            A1 = activate_a(Z1)
            A2 = activate_a(Z2)
            A3 = activate_a(Z3)
            
            For k = 1 To UBound(Z1)
                
                g.Cells(k + 1, 14) = A1(k)
                g.Cells(k + 1, 15) = A2(k)
                g.Cells(k + 1, 16) = A3(k)
                
                    If (A1(k) > 0.5) Then
                g.Cells(k + 1, 17) = 1
                    Else
                g.Cells(k + 1, 17) = 0
                    End If
            
                    If (A2(k) > 0.5) Then
                g.Cells(k + 1, 18) = 1
                     Else
                g.Cells(k + 1, 18) = 0
                    End If
             
                     If (A3(k) > 0.5) Then
                g.Cells(k + 1, 19) = 1
                    Else
                g.Cells(k + 1, 19) = 0
                    End If
        Next k
            
            
            
End Sub

Function Activate(x As Double) ' Prend en paramètre un double on peut avoir celui aussi qui prend en paramètre un array
   
   Activate = 1 / (1 + Exp(-x))
   
End Function


Function activate_a(T() As Double)
    Dim resultat() As Double
    Dim i As Integer
    For i = LBound(T) To UBound(T)
        T(i) = Activate(T(i))
    Next i
    resultat = T
    activate_a = resultat
End Function
