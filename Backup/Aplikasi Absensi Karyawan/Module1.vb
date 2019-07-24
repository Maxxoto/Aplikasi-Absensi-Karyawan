Imports System.Data.SqlClient

Module Module1

    Public CONN As SqlConnection
    Public CMD As SqlCommand
    Public DS As New DataSet
    Public DA As SqlDataAdapter
    Public RD As SqlDataReader
    Public LokasiData As String

    Sub Koneksi()
        'LokasiData = "Data Source=inspirouus;Initial Catalog=DBABSENSI;Integrated Security=True"
        LokasiData = "Data Source=user-pc;Initial Catalog=DBABSENSI;Integrated Security=True"
        CONN = New SqlConnection(LokasiData)
        If CONN.State = ConnectionState.Closed Then
            CONN.Open()
        End If
    End Sub

End Module
