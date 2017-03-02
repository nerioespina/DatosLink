Imports vb = Microsoft.VisualBasic
Imports System.Data
Imports System.Data.Odbc
Imports System.Xml
Imports MySql.Data.MySqlClient
Public Class Datos
    Public Shared connection As New MySqlConnection
    '"server=SERVERNAMEORIP;User Id=DBUSER;password=DBPASSWORD;database=DBNAME"
    Public Shared Function getInfo2(ByVal queryString As String, ByVal SC As String)

        'DEVUELVE EL RESULTADO EN UN OBJETO DATATABLE QUE PUEDE LLENAR UN DATAGRIDVIEW 
        Dim connectionString As String = SC
        Dim resultado As DataTable = New DataTable
        If connection.ConnectionString <> connectionString Then
            connection.ConnectionString = connectionString
        End If
        'Using connection As New MySqlConnection(connectionString)
        Using connection
            Dim command As New MySqlCommand(queryString, connection)
            Try
                connection.Open()
                Dim reader As MySqlDataReader = command.ExecuteReader()
                If reader.HasRows = True Then

                    resultado.Load(reader)
                    Return resultado
                End If

            Finally
                connection.Close()
            End Try
        End Using
        Return Nothing
    End Function
    Public Shared Function getInfo1(ByVal queryString As String, ByVal SC As String) As String()
        'DEVUELVE UN ARRAY DE TEXTO
        Dim connectionString As String = SC
        Dim resultado() As String = New String() {}
        'Using connection As New MySqlConnection(connectionString)
        If connection.ConnectionString <> connectionString Then
            connection.ConnectionString = connectionString
        End If
        Using connection
            Dim command As New MySqlCommand(queryString, connection)
            Dim reader As MySqlDataReader
            Try
                connection.Open()
                reader = command.ExecuteReader()
                If reader.HasRows = True Then
                    While reader.Read()
                        ReDim resultado(reader.FieldCount - 1)
                        For i As Integer = 0 To reader.FieldCount - 1
                            resultado(i) = reader.GetString(i)
                        Next
                    End While
                    Return resultado
                End If

            Finally

                connection.Close()
            End Try
            Try
                reader.Close()
            Catch ex As Exception

            End Try
        End Using
        Return Nothing
    End Function
    Public Shared Function modificar(ByVal queryString As String, SC As String) As Long
        'NO DEVUELVE INFORMACIÓN
        Dim connectionString As String = SC
        Dim affected As Long
        'Using connection As New MySqlConnection(connectionString)
        If connection.ConnectionString <> connectionString Then
            connection.ConnectionString = connectionString
        End If
        Using connection
            Dim command As New MySqlCommand(queryString, connection)
            Try
                connection.Open()
                affected = command.ExecuteNonQuery()
                connection.Close()
                Return affected
            Catch ex As Exception
                'MsgBox("Fallo en la conexión" & vbCrLf & ex.Message)
            Finally
                connection.Close()
            End Try
        End Using
        Return Nothing
    End Function
End Class
Public Class Datos1
    Public Shared connection1 As New MySqlConnection
    Public Shared connection2 As New MySqlConnection
    '"server=SERVERNAMEORIP;User Id=DBUSER;password=DBPASSWORD;database=DBNAME"
    Public Shared Function getInfo2(ByVal queryString As Object, ByVal SC As String)

        'DEVUELVE EL RESULTADO EN UN OBJETO DATATABLE QUE PUEDE LLENAR UN DATAGRIDVIEW 
        Dim connectionString As String = SC
        Dim resultado As DataTable = New DataTable
        If connection1.ConnectionString <> connectionString Then
            connection1.ConnectionString = connectionString
        End If
        'Using connection As New MySqlConnection(connectionString)
        Using connection1
            Dim command As New MySqlCommand(queryString, connection1)
            Try
                connection1.Open()
                Dim reader As MySqlDataReader = command.ExecuteReader()
                If reader.HasRows = True Then

                    resultado.Load(reader)
                    Return resultado
                End If

            Finally
                connection1.Close()
            End Try
        End Using
        Return Nothing
    End Function
    Public Shared Function getInfo1(ByVal queryString As Object, ByVal SC As String)
        'DEVUELVE UN ARRAY DE TEXTO
        Dim connectionString As String = SC
        Dim resultado() As String = New String() {}
        'Using connection As New MySqlConnection(connectionString)
        If connection2.ConnectionString <> connectionString Then
            connection2.ConnectionString = connectionString
        End If
        Using connection2
            Dim command As New MySqlCommand(queryString, connection2)

            Dim reader As MySqlDataReader
            Try
                connection2.Open()
                reader = command.ExecuteReader()
                If reader.HasRows = True Then
                    While reader.Read()
                        ReDim resultado(reader.FieldCount - 1)
                        For i As Integer = 0 To reader.FieldCount - 1
                            resultado(i) = reader.GetString(i)
                        Next
                    End While
                    Return resultado
                End If
            Finally
                reader.Close()
                connection2.Close()
            End Try
        End Using
        Return Nothing
    End Function
    Public Shared Function modificar(ByVal queryString As String, SC As String) As Long
        'NO DEVUELVE INFORMACIÓN
        Dim connectionString As String = SC
        Dim affected As Long
        'Using connection As New MySqlConnection(connectionString)
        If connection1.ConnectionString <> connectionString Then
            connection1.ConnectionString = connectionString
        End If
        Using connection1
            Dim command As New MySqlCommand(queryString, connection1)
            Try
                connection1.Open()
                affected = command.ExecuteNonQuery()
                connection1.Close()
                Return affected
            Finally
                connection1.Close()
            End Try
        End Using
        Return Nothing
    End Function
End Class