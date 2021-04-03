VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Perceptron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private L_rate As Double
Private N_iter As Double

Property Get Rate() As Double

    Rate = L_rate
    
End Property

Property Get Iter() As Double

    Iter = N_iter
    
End Property


Property Let set_Rate(a As Double)

    L_rate = a
    
End Property

Property Let set_Iter(b As Double)
    N_iter = b
End Property
Public Function P_Weights(a As Integer) As Variant

    Randomize
    Dim i As Integer
    
    Dim W_array() As Double
    ReDim W_array(1 To a)
    
    
    For i = 1 To a
    
        
         W_array(i) = Rnd()
         
    Next i
    
    P_Weights = W_array
    
End Function


Public Function Foward(tabl() As Double)

    


End Function

Public Function Backward(tabl() As Double, T() As Double)
    
    
End Function


