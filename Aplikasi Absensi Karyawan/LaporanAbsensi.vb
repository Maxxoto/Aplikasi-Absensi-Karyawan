Imports System.Data.SqlClient

Public Class LaporanAbsensi


    Private Sub LaporanAbsensi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        CMD = New SqlCommand("select distinct NIP from tblabsensi", CONN)
        RD = CMD.ExecuteReader
        Do While RD.Read
            ComboBox1.Items.Add(RD.Item("NIP"))
            ComboBox2.Items.Add(RD.Item("NIP"))
        Loop
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from tblpegawai where nip='" & ComboBox1.Text & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            TextBox1.Text = RD.Item("Nama")
            CRV.ReportSource = Nothing
            CRV.RefreshReport()
            CRV.SelectionFormula = "({tblpegawai.NIP})='" & ComboBox1.Text & "'"
            CRV.ReportSource = "lap per nip.RPT"
        Else
            MsgBox("NIP tidak terdaftar")
        End If
    End Sub

    Private Sub CRV_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CRV.Load

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        On Error Resume Next
        CRV.ReportSource = Nothing
        'CRV.RefreshReport()
        CRV.SelectionFormula = "Month({TBLAbsensi.Tanggal})=" & Month(DTP2.Text) & " and Year({TBLAbsensi.Tanggal})=" & Year(DTP3.Text)
        CRV.ReportSource = "lap per bulan.rpt"
        CRV.RefreshReport()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        On Error Resume Next
        CRV.ReportSource = Nothing
        'CRV.RefreshReport()
        CRV.SelectionFormula = "{tblabsensi.tanggal} in date ('" & DTP1.Text & "')"
        CRV.ReportSource = "lap per tanggal.rpt"
        CRV.RefreshReport()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from tblpegawai where nip='" & ComboBox2.Text & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            TextBox2.Text = RD.Item("Nama")
            CRV.ReportSource = Nothing
            CRV.RefreshReport()
            CRV.SelectionFormula = "({tblpegawai.NIP})='" & ComboBox1.Text & "'"
            CRV.ReportSource = "lap per nip.RPT"
        Else
            MsgBox("NIP tidak terdaftar")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        On Error Resume Next
        If ComboBox2.Text = "" Then
            MsgBox("NIK harus diisi dulu")
            Exit Sub
        Else
            CRV.ReportSource = Nothing
            CRV.RefreshReport()
            CRV.SelectionFormula = "({TBLtidakhadir.NIP})='" & ComboBox2.Text & "' and Month({TBLtidakhadir.Tanggal})=" & Month(DTP4.Text) & " and Year({TBLtidakhadir.Tanggal})=" & Year(DTP5.Text)
            'CRV.ReportSource = "lap karyawan per bulan.rpt"
            CRV.ReportSource = "lap tidak hadir per karyawan.rpt"
            CRV.RefreshReport()
        End If
        
    End Sub
End Class