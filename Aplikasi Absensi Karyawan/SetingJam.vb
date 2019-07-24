Imports System.Data.SqlClient

Public Class SetingJam

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Seting jam belum benar")
            Exit Sub
        End If
        Call Koneksi()
        Dim edit As String = "update tblsetingjam set jam_masuk='" & TimeValue(TextBox1.Text) & "',jam_keluar='" & TimeValue(TextBox2.Text) & "'"
        CMD = New SqlCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        MsgBox("Seting jam telah disimpan")
        
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox2.Focus()
        End If
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = ":" Or e.KeyChar = vbBack) Then e.Handled = True
    End Sub

    
    Private Sub SetingJam_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        CMD = New SqlCommand("select * from tblsetingjam", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            TextBox1.Text = TimeValue(RD.Item("Jam_Masuk"))
            TextBox2.Text = TimeValue(RD.Item("Jam_keluar"))
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class