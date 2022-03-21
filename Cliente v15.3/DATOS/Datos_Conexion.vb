Public Class Datos_Conexion
    Public Function String_Conexion() As String
        ' Return "data source = SISTEMAS\SQL2014; initial catalog = PrecitrolAdministrativo; user id = sa; password = Sistemas123"
        'Return "data source = 192.168.9.223\METROLOGIA; initial catalog = SisMetPrec; user id = sa; password = Sistemas123*"
        Return "data source = .\SRVMETROLOGIA; initial catalog = SisMetPrec; user id = sa; password = Sistemas123*"
    End Function

End Class
