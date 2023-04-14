Imports System.Data.Odbc
Imports CrystalDecisions.CrystalReports.Engine
Public Class form_report
    Dim con As OdbcConnection
    Dim dr As OdbcDataReader
    Dim da As OdbcDataAdapter
    Dim ds As DataSet
    Dim dt As DataTable
    Dim cmd As OdbcCommand
    Sub koneksi()
        con = New OdbcConnection
        con.ConnectionString = "dsn=RUMAH_MAKAN"
        con.Open()
    End Sub
    Private Sub form_report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        cmd = New OdbcCommand("select distinct id_transaksi from t_transaksi", con)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Add("pilih")
        ComboBox1.SelectedIndex = 0
        While dr.Read()
            ComboBox1.Items.Add(dr.Item("id_transaksi"))
        End While
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            CrystalReportViewer1.ReportSource = Nothing
            Exit Sub
        Else
            Dim reportku As New ReportDocument
            reportku.Load("D:\KULIAH FAHMI\prujek uas\rumahmakan\rumahmakan\transaksireport.rpt")
            reportku.SetParameterValue("trans", ComboBox1.Text)
            CrystalReportViewer1.ReportSource = reportku
            CrystalReportViewer1.Refresh()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub
End Class