Imports System.Data.SqlClient
Module Conexion
    Public con As SqlConnection
    Public cmd As SqlCommand
    Public query As String

    Sub conn()
        con = New SqlConnection
        con.ConnectionString = "Data Source=(local);Initial Catalog=MAG;User ID=sa;Password=ajb"
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
                'MsgBox("Conexion establecida")
            End If
        Catch ex As Exception
            MsgBox("Error al conectar: " & Err.Description)
        End Try
    End Sub

End Module
