
Imports System.Data.SqlClient

Public Class Pegawai

    Private Sub Pegawai_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call TampilGrid()
        Call Kosongkan()
    End Sub

    Sub Kosongkan()
        On Error Resume Next
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        PictureBox1.Load(TextBox8.Text)
        TextBox9.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        TextBox1.Focus()

    End Sub

    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select * from TBLPegawai", CONN)
        DS = New DataSet
        DA.Fill(DS, "Pegawai")
        DGV.DataSource = DS.Tables("Pegawai")
        'DGV.Columns(7).Visible = False
        DGV.ReadOnly = True
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 5
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New SqlCommand("Select * from TBLPegawai where NIP='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                On Error Resume Next
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                TextBox5.Clear()
                TextBox6.Clear()
                TextBox7.Clear()
                TextBox8.Clear()
                'PictureBox1.Image.Dispose()
                PictureBox1.Load(TextBox8.Text)
                TextBox9.Clear()
                ComboBox1.Text = ""
                ComboBox2.Text = ""
                TextBox2.Focus()
            Else
                TextBox2.Text = RD.Item("pwd")
                TextBox3.Text = RD.Item("nama")
                TextBox4.Text = RD.Item("tempat_lahir")
                DTP1.Text = RD.Item("tgl_lahir")
                TextBox5.Text = RD.Item("alamat")
                TextBox6.Text = RD.Item("telepon")
                ComboBox1.Text = RD.Item("jenis_kelamin")
                ComboBox2.Text = RD.Item("pendidikan")
                TextBox7.Text = RD.Item("jabatan")
                TextBox8.Text = RD.Item("lokasi")
                PictureBox1.Load(TextBox8.Text)
                TextBox2.Focus()
            End If
        End If
    End Sub



    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 30
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New SqlCommand("select * from tblpegawai where pwd='" & TextBox2.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                MsgBox("Coba gunakan password yang lain")
                TextBox2.Clear()
                TextBox2.Focus()
            Else
                TextBox3.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        TextBox3.MaxLength = 15
        If e.KeyChar = Chr(13) Then TextBox4.Focus()
    End Sub

    Private Sub textbox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        TextBox9.MaxLength = 15
        If e.KeyChar = Chr(13) Then DTP1.Focus()
        'If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled = True
    End Sub


    Private Sub textbox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        TextBox4.MaxLength = 15
        If e.KeyChar = Chr(13) Then TextBox6.Focus()

    End Sub

    Private Sub textbox6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        TextBox6.MaxLength = 30
        If e.KeyChar = Chr(13) Then ComboBox1.Focus()
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled = True
    End Sub

    Private Sub combobox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        ComboBox1.MaxLength = 30
        If e.KeyChar = Chr(13) Then ComboBox2.Focus()
    End Sub

    Private Sub combobox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        ComboBox1.MaxLength = 30
        If e.KeyChar = Chr(13) Then TextBox7.Focus()
    End Sub


    Private Sub textbox7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        TextBox6.MaxLength = 30
        If e.KeyChar = Chr(13) Then Button1.Focus()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or ComboBox1.Text = "" Or ComboBox2.Text = "" Or TextBox8.Text = "" Then
            MsgBox("Data belum lengkap, belum ada foto")
            OpenFileDialog1.Filter = "(*.jpg)|*.jpg|(*.bmp)|*.bmp|All files (*.*)|*.*"
            OpenFileDialog1.ShowDialog()
            PictureBox1.Load(OpenFileDialog1.FileName)
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            TextBox8.Text = OpenFileDialog1.FileName
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("Select * from TBLPegawai where NIP='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                Call Koneksi()
                Dim simpan As String = "insert into TBLPegawai values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & DTP1.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Else
                Call Koneksi()
                Dim edit As String = "update TBLPegawai set " & _
                "pwd='" & TextBox2.Text & "', " & _
                "nama='" & TextBox3.Text & "', " & _
                "tempat_lahir='" & TextBox4.Text & "', " & _
                "tgl_lahir='" & DTP1.Text & "', " & _
                "alamat='" & TextBox5.Text & "', " & _
                "telepon='" & TextBox6.Text & "', " & _
                "jenis_kelamin='" & ComboBox1.Text & "', " & _
                "pendidikan='" & ComboBox2.Text & "', " & _
                "jabatan='" & TextBox6.Text & "', " & _
                "lokasi='" & TextBox8.Text & "' where NIP='" & TextBox1.Text & "'"
                CMD = New SqlCommand(edit, CONN)
                CMD.ExecuteNonQuery()
            End If

            Call TampilGrid()
            Call Kosongkan()

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("kode Pegawai masih kosong, silakan diisi dulu")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                Dim hapus As String = "delete  from TBLPegawai where NIP='" & TextBox1.Text & "'"
                CMD = New SqlCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                Call TampilGrid()
                Call Kosongkan()
            Else
                Call Kosongkan()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub


    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBLPegawai where Nama like '%" & TextBox7.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBLPegawai where Nama like '%" & TextBox7.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS, "ketemu")
            DGV.DataSource = DS.Tables("ketemu")
            DGV.ReadOnly = True
        Else
            MsgBox("data tidak ditemukan")
        End If
    End Sub

    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(2).Value
        TextBox4.Text = DGV.Rows(e.RowIndex).Cells(3).Value
        DTP1.Text = DGV.Rows(e.RowIndex).Cells(4).Value
        TextBox5.Text = DGV.Rows(e.RowIndex).Cells(5).Value
        TextBox6.Text = DGV.Rows(e.RowIndex).Cells(6).Value
        ComboBox1.Text = DGV.Rows(e.RowIndex).Cells(7).Value
        ComboBox2.Text = DGV.Rows(e.RowIndex).Cells(8).Value
        TextBox7.Text = DGV.Rows(e.RowIndex).Cells(9).Value

        TextBox8.Text = DGV.Rows(e.RowIndex).Cells(10).Value
        PictureBox1.Load(TextBox8.Text)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        On Error Resume Next
        OpenFileDialog1.Filter = "(*.jpg)|*.jpg|(*.bmp)|*.bmp|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()
        PictureBox1.Load(OpenFileDialog1.FileName)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        TextBox8.Text = OpenFileDialog1.FileName

    End Sub

    

    Private Sub DTP1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTP1.KeyPress
        If e.KeyChar = Chr(13) Then TextBox5.Focus()
    End Sub

    Private Sub DGV_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellContentClick

    End Sub
End Class