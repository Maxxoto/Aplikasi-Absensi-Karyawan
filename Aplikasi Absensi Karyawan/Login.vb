Imports System.Data.SqlClient

Public Class Login

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then TextBox2.Focus()
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then Button1.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data login belum lengkap")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBLUSER where nama_User='" & TextBox1.Text & "' and pwd_User='" & TextBox2.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Me.Visible = False
                MenuUtama.Show()
                MenuUtama.Panel1.Text = RD.Item("kode_User")
                MenuUtama.Panel2.Text = RD.Item("nama_User")
                MenuUtama.Panel3.Text = RD.Item("status_User")
                If MenuUtama.Panel3.Text <> "ADMIN" Then
                    MenuUtama.UserToolStripMenuItem.Enabled = False
                    MenuUtama.AbsensiManualToolStripMenuItem.Enabled = False
                    MenuUtama.Button6.Enabled = False
                    MenuUtama.Button1.Enabled = False
                Else
                    MenuUtama.UserToolStripMenuItem.Enabled = True
                    MenuUtama.AbsensiManualToolStripMenuItem.Enabled = True
                    MenuUtama.Button6.Enabled = True
                    MenuUtama.Button1.Enabled = True
                End If
            Else
                MsgBox("Password salah")
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub
End Class