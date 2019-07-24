Imports System.Data.SqlClient

Public Class AbsensiManual

    Private Sub AbsensiManual_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call BuatKolom()
    End Sub

    Sub BuatKolom()
        DGV.Columns.Add("NIP", "NIP")
        DGV.Columns.Add("Nama", "Nama")
        DGV.Columns.Add("Masuk", "Jam Masuk")
        DGV.Columns.Add("Keluar", "Jam Keluar")
        DGV.Columns.Add("KetMasuk", "Keterangan1")
        DGV.Columns.Add("KetKeluar", "Keterangan2")

        DGV.Columns(0).Width = 75
        DGV.Columns(1).Width = 150
        DGV.Columns(2).Width = 75
        DGV.Columns(3).Width = 75
        DGV.Columns(2).DefaultCellStyle.Format = "hh:mm:ss"
        DGV.Columns(3).DefaultCellStyle.Format = "hh:mm:ss"
    End Sub


    Private Sub DGV_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellContentClick

    End Sub

    
    Private Sub DGV_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellEndEdit
        If e.ColumnIndex = 0 Then
            Call Koneksi()
            CMD = New SqlCommand("select * from tblpegawai where nip='" & DGV.Rows(e.RowIndex).Cells(0).Value & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                DGV.Rows(e.RowIndex).Cells(1).Value = RD.Item("nama")
                Label4.Text = DGV.RowCount - 1
            Else
                MsgBox("NIP tidak terdaftar")
            End If

        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DGV.Columns.Clear()
        Call BuatKolom()

    End Sub

    Private Sub DGV_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DGV.KeyPress
        On Error Resume Next
        If e.KeyChar = Chr(27) Then
            DGV.Rows.RemoveAt(DGV.CurrentCell.RowIndex)
            Label4.Text = DGV.RowCount - 1
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Label4.Text = "" Then
            MsgBox("Tidak ada data yang dapat disimpan")
            Exit Sub
        Else
            Call Koneksi()
            For baris As Integer = 0 To DGV.RowCount - 2
                Dim simpan As String = "Insert into tblabsensi values ('" & DTP1.Text & "','" & DGV.Rows(baris).Cells(0).Value & "','" & TimeValue(DGV.Rows(baris).Cells(2).Value) & "','" & TimeValue(DGV.Rows(baris).Cells(3).Value) & "','" & DGV.Rows(baris).Cells(4).Value & "','" & DGV.Rows(baris).Cells(5).Value & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Next
            DGV.Columns.Clear()
            Call BuatKolom()
        End If
    End Sub
End Class