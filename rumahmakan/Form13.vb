Imports System.Data.Odbc
Public Class Form13
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
        Dim sql As String = "insert into tb_nota values ('" & TextBox1.Text & "','" & ComboBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox7.Text & "','" & TextBox4.Text & "','" & TextBox8.Text & "','" & TextBox5.Text & "','" & TextBox9.Text & "','" & TextBox6.Text & "','" & TextBox10.Text & "','" & TextBox11.Text & "','" & TextBox13.Text & "','" & TextBox12.Text & "')"
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
            da = New OdbcDataAdapter("select *from tb_nota order by id_nota asc",
           con)
            dt = New DataTable
            da.Fill(dt)
            For Each row In dt.Rows
                DataGridView1.Rows.Add(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13))
            Next
            dt.Rows.Clear()
        Catch ex As Exception
            MsgBox("Menampilkan data GAGAL")
        End Try
    End Sub
    Sub NIM()
        cmd = New OdbcCommand("select id_costumer from tb_costumer", con)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox1.Items.Add(dr.Item("id_costumer"))
        Loop
    End Sub
    Sub NOM()
        cmd = New OdbcCommand("select id_transaksi from t_transaksi", con)
        dr = cmd.ExecuteReader
        ComboBox2.Items.Clear()
        Do While dr.Read
            ComboBox2.Items.Add(dr.Item("id_transaksi"))
        Loop
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        cmd = New OdbcCommand("select * from tb_costumer where id_costumer='" & ComboBox1.Text & "'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox2.Text = dr.Item("nm_costumer")

        Else
            MsgBox("id costumer tidak ada")

        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form13_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
        NIM()
        NOM()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        cmd = New OdbcCommand("select * from t_transaksi where id_transaksi='" & ComboBox2.Text & "'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox3.Text = dr.Item("makanan1")
            TextBox4.Text = dr.Item("makananmakan2")
            TextBox5.Text = dr.Item("ninuman1")
            TextBox6.Text = dr.Item("minuman2")
            TextBox7.Text = dr.Item("hargamakan1")
            TextBox8.Text = dr.Item("harga2")
            TextBox9.Text = dr.Item("hargaminum1")
            TextBox10.Text = dr.Item("hargaminum2")
            TextBox11.Text = dr.Item("total_bayar")

        Else
            MsgBox("id costumer tidak ada")

        End If
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        
    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox13.TextChanged
        Dim TOT As Integer
        Dim BY As Integer
        Dim KM As Integer

        TOT = CInt(TextBox11.Text)
        BY = CInt(TextBox13.Text)
        KM = (BY - TOT)

        TextBox12.Text = KM
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged

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
            MsgBox("Data Nota yang dihapus belum DIPILIH")
        Else
            If (MessageBox.Show("Anda yakin menghapus data dengan id_nota=" & a &
           "...?", "Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) =
           Windows.Forms.DialogResult.OK) Then
                koneksi()
                cmd = New OdbcCommand("delete from tb_nota where id_nota='" & a &
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
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = "0"
        TextBox12.Text = "0"
        TextBox13.Text = "0"

        
        ComboBox1.Text = ""
        ComboBox2.Text = ""

        TextBox1.Focus()
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

    Private Sub Panel4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel4.Click
        Form3.Show()
    End Sub

    Private Sub Panel4_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel4.Paint

    End Sub

    Private Sub Panel5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel5.Click
        Close()
    End Sub

    Private Sub Panel5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel5.Paint

    End Sub
End Class