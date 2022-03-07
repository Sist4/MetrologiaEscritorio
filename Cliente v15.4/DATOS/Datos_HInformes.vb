Imports System.Data.SqlClient

Public Class Datos_HInformes
    Dim ConexionSql As SqlConnection = Nothing
    Dim ComandoSql As SqlCommand = Nothing
    Dim query = Nothing
    Dim LectorDatos As SqlDataReader = Nothing
    Dim AdaptadorSql As SqlDataAdapter = Nothing
    Dim DatoAlmacenado As DataSet = Nothing
    Private CadenaSql As New Datos_Conexion()
    Public Function Informes_Registrador() As DataSet
        Dim Dato_Almacenado As DataSet = New DataSet()
        Dim consulta = "SELECT CodBpr,ClaBpr,ideBpr,UniDivEsc_dBpr,IdeComBpr,lugcalBpr FROM Balxpro WHERE estado  IN('Imprimir')"
        Try

            Using ConexionSql = New SqlConnection(CadenaSql.String_Conexion())
                ConexionSql.Open()
                Dim Comando_Sql As SqlCommand = New SqlCommand(consulta, ConexionSql)
                Dim Adaptador_Sql As SqlDataAdapter = New SqlDataAdapter(Comando_Sql)
                Adaptador_Sql.Fill(Dato_Almacenado)
                ConexionSql.Close()
                SqlConnection.ClearAllPools()
            End Using

        Finally

            If ConexionSql IsNot Nothing AndAlso ConexionSql.State <> ConnectionState.Closed Then
                ConexionSql.Close()
            End If
        End Try

        Return Dato_Almacenado
    End Function

    Public Function Masa_Conveccioanla(proyecto As String, Cargas As String) As DataSet
        Dim Dato_Almacenado As DataSet = New DataSet()
        Dim consulta = ""
        Try
            consulta = "select  
                    CASE WHEN sum(N1) IS NULL THEN 0 ELSE sum(N1) END + 
                    CASE WHEN sum(N2) IS NULL THEN 0 ELSE sum(N2) END +
                    CASE WHEN sum(N2A) IS NULL THEN 0 ELSE sum(N2A) END +
                    CASE WHEN sum(N5) IS NULL THEN 0 ELSE sum(N5) END +
                    CASE WHEN sum(N10) IS NULL THEN 0 ELSE sum(N10) END +
                    CASE WHEN sum(N20) IS NULL THEN 0 ELSE sum(N20) END +
                    CASE WHEN sum(N20A) IS NULL THEN 0 ELSE sum(N20A) END +
                    CASE WHEN sum(N50) IS NULL THEN 0 ELSE sum(N50) END +
                    CASE WHEN sum(N100) IS NULL THEN 0 ELSE sum(N100) END +
                    CASE WHEN sum(N200) IS NULL THEN 0 ELSE sum(N200) END +
                    CASE WHEN sum(N200A) IS NULL THEN 0 ELSE sum(N200A) END +
                    CASE WHEN sum(N500) IS NULL THEN 0 ELSE sum(N500) END +
                    CASE WHEN sum(N1000) IS NULL THEN 0 ELSE sum(N1000) END +
                    CASE WHEN sum(N2000) IS NULL THEN 0 ELSE sum(N2000) END +
                    CASE WHEN sum(N2000A) IS NULL THEN 0 ELSE sum(N2000A) END +
                    CASE WHEN sum(N5000) IS NULL THEN 0 ELSE sum(N5000) END+
                    CASE WHEN sum(N10000) IS NULL THEN 0 ELSE sum(N10000) END +
                    CASE WHEN sum(N20000) IS NULL THEN 0 ELSE sum(N20000) END +
                    CASE WHEN sum(N1000000) IS NULL THEN 0 ELSE sum(N1000000) END 

                    from V_MasaConvencional where IdeComBpr='" & proyecto & "'  group by IdeComBpr,TipPxp ORDER BY CONVERT(int,REPLACE(REPLACE(TipPxp,'C',''),'+',''))"

            Using ConexionSql = New SqlConnection(CadenaSql.String_Conexion())

                Dim Comando_Sql As SqlCommand = New SqlCommand(consulta, ConexionSql)
                Dim Adaptador_Sql As SqlDataAdapter = New SqlDataAdapter(Comando_Sql)
                Adaptador_Sql.Fill(Dato_Almacenado)
                ConexionSql.Close()
                SqlConnection.ClearAllPools()
            End Using

        Finally

            If ConexionSql IsNot Nothing AndAlso ConexionSql.State <> ConnectionState.Closed Then
                ConexionSql.Close()
            End If
        End Try

        Return Dato_Almacenado
    End Function
End Class
