Public Class LaporanMaster

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        CRV.ReportSource = Nothing
        CRV.RefreshReport()
        CRV.ReportSource = "LAP USER.RPT"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        CRV.ReportSource = Nothing
        CRV.RefreshReport()
        CRV.ReportSource = "LAP PEGAWAI.RPT"
    End Sub
End Class