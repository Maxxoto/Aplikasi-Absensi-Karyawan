Imports System.Data.SqlClient

Public Class AbsensiMasuk

    Private Sub AbsensiMasuk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Label5.Text = Format(Today, "MM/dd/yyyy")


        CMD = New SqlCommand("select * from tblsetingjam", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Label1.Text = Format(RD.Item("jam_masuk"), "hh:mm")
        End If
    End Sub

    Sub Kosongkan()
        On Error Resume Next
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Text = "-"
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
                'TextBox2.Text = RD.Item("NIP")
                

                Call Koneksi()
                CMD = New SqlCommand("select * from tblabsensi where tanggal='" & Label5.Text & "' and NIP ='" & TextBox1.Text & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If RD.HasRows Then
                    MsgBox("Absensi masuk anda hari ini sudah tersimpan")
                    Call Kosongkan()
                    Exit Sub
                End If
                
                If TimeValue(Label2.Text) > TimeValue(Label1.Text) And TextBox5.Text = "-" Then
                    MsgBox("anda masuk terlambat, mohon mengisi alasan keterlambatan")
                    TextBox5.Focus()
                    Exit Sub
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

        If TimeValue(Label2.Text) > TimeValue(Label1.Text) And TextBox5.Text = "-" Then
            MsgBox("anda masuk terlambat, mohon mengisi alasan keterlambatan")
            TextBox5.Focus()
            Exit Sub
        End If
            
            Call Koneksi()
            Dim simpan As String = "insert into tblabsensi values ('" & Label5.Text & "','" & TextBox1.Text & "','" & Label2.Text & "',0,'" & TextBox5.Text & "','-')"
            CMD = New SqlCommand(simpan, CONN)
            CMD.ExecuteNonQuery()
            Call Kosongkan()


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar = Chr(13) Then Button1.Focus()
    End Sub

    
End Class