VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeneradorPseudoaleatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim multiplicador As Long
Dim constante_aditiva As Long
Dim modulo As Long
Dim semilla As Long
Dim cantidad As Long
Dim numeros_generados() As Double

Public Sub IngresarValores(ByVal a As Long, ByVal c As Long, ByVal x0 As Long, ByVal m As Long, ByVal N As Long)
    multiplicador = a
    constante_aditiva = c
    semilla = x0
    modulo = m
    cantidad = N
    ReDim numeros_generados(N) As Double
End Sub

Private Function valor(ByVal xi As Double) As Double
    valor = ((multiplicador * xi) + constante_aditiva) Mod modulo
End Function

Private Sub xn()
    Dim xi As Double
    xi = valor(semilla)
    almacenar xi, 0
    For i = 1 To cantidad - 1
        xi = valor(xi)
        almacenar xi, i
    Next
End Sub

Public Function GenerarNumeros() As Double()
    xn
    GenerarNumeros = numeros_generados
End Function


Private Sub almacenar(ByVal xi As Double, ByVal i As Integer)
    Dim rn As Double
    rn = xi / modulo
    rn = Round(rn, 5)
    numeros_generados(i) = rn
End Sub

Public Function GenerarDefault() As Double()
    ReDim numeros_generados(40) As Double
    numeros_generados(0) = 0.35
    numeros_generados(1) = 0.87
    numeros_generados(2) = 0.12
    numeros_generados(3) = 0.64
    numeros_generados(4) = 0.03
    numeros_generados(5) = 0.92
    numeros_generados(6) = 0.75
    numeros_generados(7) = 0.24
    numeros_generados(8) = 0.02
    numeros_generados(9) = 0.41
    numeros_generados(10) = 0.57
    numeros_generados(11) = 0.36
    numeros_generados(12) = 0.49
    numeros_generados(13) = 0.11
    numeros_generados(14) = 0.8
    numeros_generados(15) = 0.73
    numeros_generados(16) = 0.23
    numeros_generados(17) = 0.1
    numeros_generados(18) = 0.61
    numeros_generados(19) = 0.06
    numeros_generados(20) = 0.09
    numeros_generados(21) = 0.26
    numeros_generados(22) = 0.42
    numeros_generados(23) = 0.97
    numeros_generados(24) = 0.7
    numeros_generados(25) = 0.38
    numeros_generados(26) = 0.16
    numeros_generados(27) = 0.04
    numeros_generados(28) = 0.62
    numeros_generados(29) = 0.89
    numeros_generados(30) = 0.54
    numeros_generados(31) = 0.56
    numeros_generados(32) = 0.91
    numeros_generados(33) = 0.32
    numeros_generados(34) = 0.72
    numeros_generados(35) = 0.09
    numeros_generados(36) = 0.15
    numeros_generados(37) = 0.47
    numeros_generados(38) = 0.68
    numeros_generados(39) = 0.2
    GenerarDefault = numeros_generados
    
End Function
    
