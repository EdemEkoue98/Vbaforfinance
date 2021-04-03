Attribute VB_Name = "Module2"
Option Explicit

Sub Perceptron_G() 'perceptron linéaire de type separation de 2 groupes
    Dim b As Integer
    Dim Wei(1 To 1763) As Double
    Dim N_iter As Integer
    Dim a As Double ' Learning rate
    Dim Z As Double ' ce que retourne la sum et c'était le Y
    Dim Y As Double
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim M As Integer
    Dim f As Worksheet
    Set f = ThisWorkbook.Sheets("Feuil1")
    '/------------données initiales---------------------------------------/*
    
        For M = 1 To 1763
    
            Wei(M) = Rnd()
            
            f.Cells(M + 1, 12) = Wei(M)
            
        Next M
 
    
        a = Range("I2") 'Learning rate
        b = 0 'valeur du biais
    
        N_iter = Range("J2")  'Nombre d'iteration max
    
        Sheets("Feuil1").Select
    
    

    '/------------boucle sur les données---------------------------------/*
        For k = 2 To 1412 '8 lignes d'apprentissage
        
                Y = f.Cells(k, 3) ' on selectionne  les features
        
    '/-----------vecteur apprentissage-----------------------------------/*
                Z = (f.Cells(k, 2) * Wei(k - 1)) + b 'on Calcule les les Zi = wei(k)*Xi ou Xi = returns
                f.Cells(k, 15) = Z
                
    '/----------'Fonction seuil, renvoie 1 si y > 0 sinon -1--------/*
                If Z > 0 Then
                    Z = 1
                Else
                    Z = 0
                End If
                   
    '/-----------test pour mise à jour des poids------------------------/*
                    If Z <> Y Then
                    
                        Wei(k - 1) = Wei(k - 1) + (f.Cells(k, 2) * a * Y)
                        
                        b = b + Y * a
                        
                       '/ j = j + 1
                       
                       f.Cells(k, 13) = Wei(k - 1)
                        f.Cells(k, 14) = b
                        
                    End If
                    
                    
                        
        Next k
   '/ Loop
    '/-----------Test pour convergence ou pas------------------------------/*
   
    
    
    
    f.Cells(k, 13) = Wei(k)
    f.Cells(k, 14) = b
    
    
    
    
    
End Sub


Function Weights()

    Dim W(1 To 1763) As Double
    Dim per As Workbook
    Dim go As Worksheet
    Dim i As Integer
    
    
    Set per = ThisWorkbook
    Set go = per.Sheets("Feuil1")
    
    For i = 1 To 1763
       W(i) = Rnd()
    Next i
 
    Weights = W
 
End Function

