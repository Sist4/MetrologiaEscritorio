Imports System.Data
Imports DATOS

Public Class Negocio_HInformes

    Dim Informes As New Datos_HInformes()
    Public Function Informes_Registrador() As DataSet
        Return Informes.Informes_Registrador()
    End Function
    Public Function Masa_Conveccioanla(proyecto As String, Cargas As String) As DataSet
        Return Informes.Masa_Conveccioanla(proyecto, Cargas)
    End Function
End Class
