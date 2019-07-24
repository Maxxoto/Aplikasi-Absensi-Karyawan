Imports System.Data.SqlClient
Public Class AbsensiKeluar

    Private Sub AbsenKeluar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Label5.Text = Format(Today, "MM/dd/yyyy")
        CMD = New SqlCommand("select * from tblsetingjam", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            'Label1.Text = Format(RD.Item("jam_Keluar"), "hh:mm")
            Label1.Text = TimeValue(RD.Item("jam_Keluar"))
        End If

    End Sub

    Sub Kosongkan()
        'On Error Resume Next
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        PictureBox1.Load(TextBox4.Text)
        TextBox1.Focus()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label2.Text = TimeOfDay
    End Sub


    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New SqlCommand("select * from tblpegawai where NIP='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                TextBox2.Text = RD.Item("nama")
                TextBox2.Enabled = False
                TextBox4.Text = RD.Item("lokasi")
                PictureBox1.Load(TextBox4.Text)
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                TextBox3.Focus()
            Else
                MsgBox("NIP karyawan tidak terdaftar")
                Call Kosongkan()
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New SqlCommand("select * from tblpegawai where NIP='" & TextBox1.Text & "' and pwd= '" & TextBox3.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                If TimeValue(Label2.Text) < TimeValue(Label1.Text) Then
                    MsgBox("anda tidak pulang lebih cepat")
                    Me.Close()
                    Exit Sub
                Else
                    Button1.Enabled = True
                    Button1.Focus()
                End If

                'Button1.Focus()
            Else
                MsgBox("password salah")
                TextBox3.Clear()
                TextBox3.Focus()
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data belum lengkap")
            If TextBox1.Text = "" Then TextBox1.Focus()
            If TextBox3.Text = "" Then TextBox3.Focus()
            Exit Sub
        End If


        If TimeValue(Label2.Text) < TimeValue(Label1.Text) Then
            MsgBox("anda tidak boleh pulang lebih cepat")
            Me.Close()
            Exit Sub
        Else
            Button1.Enabled = True
            Button1.Focus()
        End If

        Call Koneksi()
        Dim simpan As String = "update tblabsensi set keluar='" & Label2.Text & "',ket_keluar='' where tanggal='" & Label5.Text & "' and nip='" & TextBox1.Text & "'"
        CMD = New SqlCommand(simpan, CONN)
        CMD.ExecuteNonQuery()
        Call Kosongkan()


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then Button1.Focus()
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.TextChanged
        'If TimeValue(Label1.Text) = TimeValue(Label2.Text) Then
        '    Button1.Enabled = False
        'Else
        '    Button1.Enabled = True
        'End If
    End Sub
End Class