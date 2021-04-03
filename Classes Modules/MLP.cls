VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MP() As Perceptron
Private Number_perceptron As Integer

Property Get Number()

        Number = Number_perceptron
        
End Property
Property Get M_P()

    M_P = MP

End Property


Property Let set_MP(n As Integer)
    
    Number_perceptron = n
    ReDim MP(1 To n) As Perceptron

End Property

Public Function Weightss(a As Integer)

    Dim W() As Double
    ReDim W(1 To a, 1 To a) As Double
    Dim i As Integer
    For i = 1 To UBound(MP)
        W(i, j) = MP(i).P_Weights(a)(i, j)
    Next i
    
    Weightss = W

    
End Function

Sub Add_Perceptron(P As Perceptron, i As Integer)

    MP(i) = P
    
End Sub
