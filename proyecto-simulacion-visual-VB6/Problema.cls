VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Problema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Contaminantes(7) As String
Dim Probabilidad_Contaminacion(7) As Single

Dim Analisis_Sangre(5) As String
Dim Probabilidad_Analisis_Sangre(5) As Single

Public Function GetContaminantes(ByVal indice As Integer) As String
    GetContaminantes = Contaminantes(indice)
End Function

Public Function GetProbabilidad_Contaminacion(ByVal indice As Integer) As String
    GetProbabilidad_Contaminacion = Probabilidad_Contaminacion(indice)
End Function

Public Function GetAnalisis_Sangre(ByVal indice As Integer) As String
    GetAnalisis_Sangre = Analisis_Sangre(indice)
End Function

Public Function GetProbabilidad_Analisis_Sangre(ByVal indice As Integer) As String
    GetProbabilidad_Analisis_Sangre = Probabilidad_Analisis_Sangre(indice)
End Function

Private Sub SetParameters()
    Contaminantes(0) = "Sustancias coloidales (SC)"
    Contaminantes(1) = "Exceso de mercurio (EM)"
    Contaminantes(2) = "Residuos petroquimicos (RP)"
    Contaminantes(3) = "Sulfatos (SF)"
    Contaminantes(4) = "Acido clorhídrico (AC)"
    Contaminantes(5) = "Fosfatos (FF)"
    Contaminantes(6) = "Óxidos (OX)"
    
    Probabilidad_Contaminacion(0) = 0.05
    Probabilidad_Contaminacion(1) = 0.1
    Probabilidad_Contaminacion(2) = 0.25
    Probabilidad_Contaminacion(3) = 0.15
    Probabilidad_Contaminacion(4) = 0.12
    Probabilidad_Contaminacion(5) = 0.16
    Probabilidad_Contaminacion(6) = 0.17
    
    Analisis_Sangre(0) = "Alto grado de acidez (AC)"
    Analisis_Sangre(1) = "Estado de anemia aguda (AA)"
    Analisis_Sangre(2) = "Estado en rango normal (RN)"
    Analisis_Sangre(3) = "Exceso de glucosa (EG)"
    Analisis_Sangre(4) = "Alto grado de alcalinidad (GA)"
    
    Probabilidad_Analisis_Sangre(0) = 0.18
    Probabilidad_Analisis_Sangre(1) = 0.08
    Probabilidad_Analisis_Sangre(2) = 0.35
    Probabilidad_Analisis_Sangre(3) = 0.17
    Probabilidad_Analisis_Sangre(4) = 0.22
End Sub

Private Sub Class_Initialize()
SetParameters
End Sub

