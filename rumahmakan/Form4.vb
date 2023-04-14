Imports System.Data.Odbc
Public Class Form4
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
    Sub simpan()
        koneksi()
        Dim sql As String = "insert into tb_suplier values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox5.Text & "')"
        cmd = New OdbcCommand(sql, con)
        cmd.ExecuteNonQuery()
        Try
            MsgBox("Menyimpan data BERHASIL", vbInformation, "INFORMASI")
        Catch ex As Exception
            MsgBox("Menyimpan data GAGAL", vbInformation, "PERINGATAN")
        End Try
    End Sub
    Sub tampil()
        DataGridView1.Rows.Clear()
        Try
            koneksi()
            da = New OdbcDataAdapter("select *from tb_suplier order by id_suplier asc",
           con)
            dt = New DataTable
            da.Fill(dt)
            For Each row In dt.Rows
                DataGridView1.Rows.Add(row(0), row(1), row(2), row(3))
            Next
            dt.Rows.Clear()
        Catch ex As Exception
            MsgBox("Menampilkan data GAGAL")
        End Try
    End Sub
    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim a As String = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        If a = "" Then
            MsgBox("Data suplier yang dihapus belum DIPILIH")
        Else
            If (MessageBox.Show("Anda yakin menghapus data dengan id_suplier=" & a &
           "...?", "Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) =
           Windows.Forms.DialogResult.OK) Then
                koneksi()
                cmd = New OdbcCommand("delete from tb_suplier where id_suplier='" & a & "'", con)
                cmd.ExecuteNonQuery()
                MsgBox("Menghapus data BERHASIL", vbInformation, "INFORMASI")
                con.Close()
                tampil()
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        simpan()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        tampil()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox5.Focus()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Close()

    End Sub

    Private Sub Panel6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel6.Click
        Form1.Show()
    End Sub

    Private Sub Panel6_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel6.Paint

    End Sub

    Private Sub Panel7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel7.Click
        Form2.Show()
    End Sub

    Private Sub Panel7_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel7.Paint

    End Sub

    Private Sub Panel8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel8.Click
        Form5.Show()
    End Sub

    Private Sub Panel8_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel8.Paint

    End Sub

    Private Sub OvalShape1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OvalShape1.Click

    End Sub

    Private Sub Panel5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel5.Click
        Close()
    End Sub

    Private Sub Panel5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel5.Paint

    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub
End Class