Attribute VB_Name = "Tangent_init"

Sub main2()

    Perceptron_2
    loss_2
    Prediction_2
 
End Sub

Sub Perceptron_2()

    Dim per As New Perceptron
    Dim a As Double
    Dim b As Double
    Dim k As Integer
    Dim i As Integer
    Dim j As Integer
    Dim f As Worksheet
    Set f = ThisWorkbook.Sheets("Feuil1")
    Dim g As Worksheet
    Set g = ThisWorkbook.Sheets("Feuil3")
    Dim P As Integer ' destiné à représenter 80% de la data
    Dim LR As Double
    LR = 0.1
    Dim NI As Integer
    NI = 20
    b = 0 ' biais
    per.set_Rate = LR ' on set le Learning rate
    per.set_Iter = NI ' on set le nombre d'iterations
    Debug.Print (per.Rate())
    Debug.Print (per.Iter())
    
    ' Preparation of the table of weights
    
    Dim W1() As Double
    W1 = per.P_Weights(3)
    ' remplissage weights feuill2
    
    For k = 1 To 3
        g.Cells(k + 1, 4) = W1(k)
    Next k
    
    ' Preparation of the Xi
    
    Dim X1() As Double
    ReDim X1(1 To 1742)                  'SMA standard deviation
    Dim X2() As Double  'UP
    ReDim X2(1 To 1742)
    Dim X3() As Double  'LB
    ReDim X3(1 To 1742)
    
    
    For k = 21 To 1762
        
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
    ReDim Z1(1 To 1742)
    
    For k = 1 To UBound(X1)
    
        Z1(k) = W1(1) * X1(k) + W1(2) * X2(k) + W1(3) * X3(k) + b
        
        g.Cells(k + 1, 5) = Z1(k)
        
        If (Z1(k) > 0) Then
            g.Cells(k + 1, 8) = 1
        Else
            g.Cells(k + 1, 8) = 0
        End If
        
    Next k
    
        ' Valeurs prédites avec fonctions d'activations Y^
            
            'Preparation of Ai
            Dim A1() As Double
            ReDim A1(1 To 1742)
            'Dim A2() As Double
            'ReDim A2(1 To 1743)
            'Dim A3() As Double
            'ReDim A3(1 To 1743)
            
            ' calcul des Ai
            
            A1 = activate_Tangent(Z1)
            
            For k = 1 To UBound(Z1)
                
                g.Cells(k + 1, 6) = A1(k)
                
            Next k
        
    
    ' gradient descent algorithm
        'Remplissage feuille 2 avec la target Y dans feuille1 on en prend que 1743
        For k = 2 To 1743
            g.Cells(k, 7) = f.Cells(k + 19, 2)
            g.Cells(k, 9) = f.Cells(k + 19, 3)
            
        Next k
        
        
            
        
    
    
End Sub

Sub loss_2()
    Dim g As Worksheet
    Set g = ThisWorkbook.Sheets("Feuil3")
       
        Dim dw1 As Double
        Dim dw2 As Double
        Dim dw3 As Double
        Dim db As Double
        
        Dim W1 As Double
        Dim W2 As Double
        Dim W3 As Double
        Dim k As Integer
        Dim LR As Double
        LR = 0.1
        Dim b As Double
        
        W1 = g.Cells(2, 4)
        W2 = g.Cells(3, 4)
        W3 = g.Cells(4, 4)
        Dim J1 As Double
        For k = 2 To 1743
        
            J1 = -(0.5) * ((1 - g.Cells(k, 7)) * Log(1 - g.Cells(k, 6)) + (1 - g.Cells(k, 7)) * Log(1 - g.Cells(k, 6))) + Log(2)
            
            g.Cells(k, 10) = J1
            
            dw1 = dw1 + (g.Cells(k, 6) - g.Cells(k, 7)) * g.Cells(k, 1)
            g.Cells(k, 11) = dw1
            
            dw2 = dw2 + (g.Cells(k, 6) - g.Cells(k, 7)) * g.Cells(k, 2)
            g.Cells(k, 12) = dw2
            
            dw3 = dw3 + (g.Cells(k, 6) - g.Cells(k, 7)) * g.Cells(k, 3)
            g.Cells(k, 13) = dw3
            
            db = db + (g.Cells(k, 6) - g.Cells(k, 7))
            g.Cells(k, 14) = db
            
        Next k
        
        
        g.Cells(2, 18) = J1 / 1742
        dw1 = dw1 / 1742
        dw2 = dw2 / 1742
        dw3 = dw3 / 1742
        'Debug.Print (db)
        db = db / 1742
        W1 = W1 - LR * dw1
        W2 = W2 - LR * dw2
        W3 = W3 - LR * dw3
        b = b - LR * db
        g.Cells(2, 19) = b
        g.Cells(2, 15) = W1
        g.Cells(2, 16) = W2
        g.Cells(2, 17) = W3
        

End Sub

Function Activate_T(x As Double) ' Prend en paramètre un double on peut avoir celui aussi qui prend en paramètre un array
   
   Activate_T = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
   
End Function


Function activate_Tangent(T() As Double)
    Dim resultat() As Double
    Dim i As Integer
    For i = LBound(T) To UBound(T)
        T(i) = Activate_T(T(i))
    Next i
    resultat = T
    activate_Tangent = resultat
End Function


Sub Prediction_2()

    Dim a As Double
    Dim b As Double
    Dim k As Integer
    Dim i As Integer
    Dim j As Integer
    Dim f As Worksheet
    Set f = ThisWorkbook.Sheets("Feuil1")
    Dim g As Worksheet
    Set g = ThisWorkbook.Sheets("Feuil3")
    Dim P As Integer ' destiné à représenter 20% de la data
    LR = 0.1
    NI = 20
    P = 350
    b = g.Cells(2, 19) ' biais adapté
    'per.set_Rate = LR ' on set le Learning rate
    'per.set_Iter = NI ' on set le nombre d'iterations
    'Debug.Print (per.Rate())
    'Debug.Print (per.Iter())
    
    ' Preparation of the table of weights
    
    Dim W1(1 To 3) As Double
    
    ' remplissage weights feuill2 mais des weights appris
    
    W1(1) = g.Cells(2, 15)
    W1(2) = g.Cells(2, 16)
    W1(3) = g.Cells(2, 17)
    
    g.Cells(2, 26) = g.Cells(2, 15)
    g.Cells(2, 27) = g.Cells(2, 16)
    g.Cells(2, 28) = g.Cells(2, 17)
    
    ' Preparation of the Xi
    
    Dim X1() As Double
    ReDim X1(1 To P)                  'SMA standard deviation
    Dim X2() As Double  'UP
    ReDim X2(1 To P)
    Dim X3() As Double  'LB
    ReDim X3(1 To P)
    
    
    For k = 1395 To 1743
         
        X1(k - 1394) = g.Cells(k, 1)
        X2(k - 1394) = g.Cells(k, 2)
        X3(k - 1394) = g.Cells(k, 3)
        g.Cells(k - 1393, 32) = f.Cells(k, 3)
    
    Next k
    
    For k = 1 To P
    
        g.Cells(k + 1, 23) = X1(k)
        g.Cells(k + 1, 24) = X2(k)
        g.Cells(k + 1, 25) = X3(k)
    
    Next k
        
        ' Preparation of the Z
    
    Dim Z1() As Double
    ReDim Z1(1 To P)
    
    For k = 1 To P
    
        Z1(k) = W1(1) * X1(k) + W1(2) * X2(k) + W1(3) * X3(k) + b
        
        g.Cells(k + 1, 29) = Z1(k)
        
        If (Z1(k) > 0) Then
            g.Cells(k + 1, 30) = 1
        Else
            g.Cells(k + 1, 30) = 0
        End If
        
    Next k
    
    
     'Preparation of Ai
            Dim A1() As Double
            ReDim A1(1 To UBound(X1))
          
            ' calcul des Ai
            
            A1 = activate_Tangent(Z1)
            
            For k = 1 To UBound(Z1)
                
                g.Cells(k + 1, 31) = A1(k)
                
            Next k
        
    
    
End Sub


