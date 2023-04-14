Imports System.Data.Odbc
Public Class Form3
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
        Dim sql As String = "insert into tb_minuman values('" & TextBox4.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox5.Text & "')"
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
            da = New OdbcDataAdapter("select *from tb_minuman order by id_minuman asc", con)
            dt = New DataTable
            da.Fill(dt)
            For Each row In dt.Rows
                DataGridView1.Rows.Add(row(0), row(1), row(2), row(3), row(4))
            Next
            dt.Rows.Clear()
        Catch ex As Exception
            MsgBox("Menampilkan data GAGAL")
        End Try
    End Sub
        Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
        End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        simpan()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        tampil()
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim a As String = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        If a = "" Then
            MsgBox("Data Customer yang dihapus belum DIPILIH")
        Else
            If (MessageBox.Show("Anda yakin menghapus data dengan id_minuman=" & a &
           "...?", "Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) =
           Windows.Forms.DialogResult.OK) Then
                koneksi()
                cmd = New OdbcCommand("delete from tb_minuman where id_minuman='" & a &
               "'", con)
                cmd.ExecuteNonQuery()
                MsgBox("Menghapus data BERHASIL", vbInformation, "INFORMASI")
                con.Close()
                tampil()
            End If
        End If
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox4.Focus()
    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

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
        Form14.Show()
    End Sub

    Private Sub Panel5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel5.Click
        Close()
    End Sub

    Private Sub Panel5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel5.Paint

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Form2.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Close()

    End Sub
End Class