Imports System.Data.SqlClient

Public Class TidakHadir

    Sub Kosongkan()
        On Error Resume Next
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Text = ""
        Label4.Text = ""

        PictureBox1.Load(Label4.Text)
        TextBox1.Focus()
        'PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
    End Sub

    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("Select * from TBLTidakHadir", CONN)
        DS = New DataSet
        DA.Fill(DS, "Ketemu")
        DGV.DataSource = DS.Tables("ketemu")
        DGV.ReadOnly = True
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New SqlCommand("select * from tblpegawai where NIP='" & TextBox1.Text & "' ", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                TextBox2.Text = RD.Item("nama")
                TextBox3.Text = RD.Item("jabatan")
                Label4.Text = RD.Item("lokasi")
                PictureBox1.Load(Label4.Text)
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                Call Koneksi()
                CMD = New SqlCommand("select * from TBLTidakHadir where NIP='" & TextBox1.Text & "' and tanggal='" & DTP1.Text & "' ", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If RD.HasRows Then
                    ComboBox1.Text = RD.Item("Keterangan")
                    ComboBox1.Focus()
                End If
                ComboBox1.Focus()
            Else
                MsgBox("NIP tidak terdaftar")
                Call Kosongkan()
            End If
        End If

    End Sub

    Private Sub DTP1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTP1.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox1.Focus()
        End If
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Data belum lengkap")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBLTidakHadir where NIP='" & TextBox1.Text & "'  and tanggal='" & DTP1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()
                Dim edit As String = "update TBLTidakHadir set keterangan='" & ComboBox1.Text & "' where tanggal='" & DTP1.Text & "' and nip='" & TextBox1.Text & "'"
                CMD = New SqlCommand(edit, CONN)
                CMD.ExecuteNonQuery()

            Else
                Call Koneksi()
                Dim simpan As String = "insert into TBLTidakHadir values ('" & DTP1.Text & "','" & TextBox1.Text & "','" & ComboBox1.Text & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()

            End If

            Call Kosongkan()
            Call TampilGrid()

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("NIK harus diisi dulu")
            TextBox1.Focus()
            Exit Sub
        Else
            Call Koneksi()

            If MessageBox.Show("Yakin data ini akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Dim hapus As String = "delete from TBLTidakHadir where nip='" & TextBox1.Text & "' and tanggal='" & DTP1.Text & "'"
                CMD = New SqlCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                Call Kosongkan()
                Call TampilGrid()
            Else
                Call Kosongkan()
                Call TampilGrid()
            End If
        End If
    End Sub

    Private Sub TidakHadir_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call TampilGrid()
        Call Kosongkan()

    End Sub
End Class