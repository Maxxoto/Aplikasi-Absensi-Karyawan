Imports System.Data.SqlClient

Module Module1

    Public CONN As SqlConnection
    Public CMD As SqlCommand
    Public DS As DataSet
    Public DA As SqlDataAdapter
    Public RD As SqlDataReader
    Public LokasiData As String

    Sub Koneksi()
        LokasiData = "Server=MOSTODANI-PC\SQLExpress;AttachDbFilename=H:\Aplikasi Absensi Karyawan\DBAbsensi_Data.mdf;Database=DBAbsensi_Data;Trusted_Connection=Yes;User ID=sa;Password=1234"
        CONN = New SqlConnection(LokasiData)
            If CONN.State = ConnectionState.Closed Then
                CONN.Open()
        End If

    End Sub
End Module
