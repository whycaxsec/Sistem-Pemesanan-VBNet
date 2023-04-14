Imports System.Data.Odbc
Public Class Form5
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
        Dim sql As String = "insert into tb_transksi values('" & TextBox9.Text & "','" & ComboBox2.Text & "','" & ComboBox4.Text & "','" & ComboBox3.Text & "','" & TextBox10.Text & "','" & TextBox11.Text & "','" & tbayar.Text & "','" & dt1.Value & "','" & dt2.Value & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')"
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
            da = New OdbcDataAdapter("select *from tb_transksi order by id_trnsaksi asc",
           con)
            dt = New DataTable
            da.Fill(dt)
            For Each row In dt.Rows
                DataGridView1.Rows.Add(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11))
            Next
            dt.Rows.Clear()
        Catch ex As Exception
            MsgBox("Menampilkan data GAGAL")
        End Try
    End Sub
    Sub NIM()
        cmd = New OdbcCommand("select id_costumer from tb_costumer", con)
        dr = cmd.ExecuteReader
        ComboBox2.Items.Clear()
        Do While dr.Read
            ComboBox2.Items.Add(dr.Item("id_costumer"))
        Loop
    End Sub
    Sub NOM()
        cmd = New OdbcCommand("select id_makanan from tb_makanan", con)
        dr = cmd.ExecuteReader
        ComboBox4.Items.Clear()
        Do While dr.Read
            ComboBox4.Items.Add(dr.Item("id_makanan"))
        Loop
    End Sub
    Sub NEM()
        cmd = New OdbcCommand("select id_minuman from tb_minuman", con)
        dr = cmd.ExecuteReader
        ComboBox3.Items.Clear()
        Do While dr.Read
            ComboBox3.Items.Add(dr.Item("id_minuman"))
        Loop
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub
    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        cmd = New OdbcCommand("select * from tb_costumer where id_costumer='" & ComboBox2.Text & "'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox6.Text = dr.Item("alm_costumer")
            TextBox7.Text = dr.Item("kota")
            TextBox8.Text = dr.Item("no_telepon")
        Else
            MsgBox("id minuman tidak ada")

        End If
    End Sub

    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
        NIM()
        NOM()
        NEM()



    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        cmd = New OdbcCommand("select * from tb_makanan where id_makanan='" & ComboBox4.Text & "'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox1.Text = dr.Item("harga_makanan")
            TextBox3.Text = dr.Item("nm_makanan")
            
        Else
            MsgBox("id makanan tidak ada")

        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        cmd = New OdbcCommand("select * from tb_minuman where id_minuman='" & ComboBox3.Text & "'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox2.Text = dr.Item("harga_minuman")
            TextBox12.Text = dr.Item("nm_minuman")
        Else
            MsgBox("id minuman tidak ada")

        End If
        
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        Dim minum As Integer
        Dim makan As Integer
        Dim mbeli As Integer
        Dim nbeli As Integer
        Dim total As Integer

        minum = CInt(TextBox1.Text)
        makan = CInt(TextBox2.Text)
        mbeli = CInt(TextBox10.Text)
        nbeli = CInt(TextBox11.Text)
        total = ((minum * mbeli) + (makan * nbeli))
        tbayar.Text = total

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        simpan()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        tampil()
    End Sub

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim a As String = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        If a = "" Then
            MsgBox("Data Customer yang dihapus belum DIPILIH")
        Else
            If (MessageBox.Show("Anda yakin menghapus data dengan id_trnsaksi=" & a &
           "...?", "Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) =
           Windows.Forms.DialogResult.OK) Then
                koneksi()
                cmd = New OdbcCommand("delete from tb_transksi where id_trnsaksi='" & a &
               "'", con)
                cmd.ExecuteNonQuery()
                MsgBox("Menghapus data BERHASIL", vbInformation, "INFORMASI")
                con.Close()
                tampil()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = "0"
        TextBox2.Text = "0"
        tbayar.Text = "0"
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = "0"
        TextBox11.Text = "0"
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        TextBox12.Text = ""
        TextBox3.Text = ""




        TextBox9.Focus()
    End Sub

    Private Sub Panel7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel7.Click
        Form2.Show()
    End Sub

    Private Sub Panel7_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel7.Paint

    End Sub

    Private Sub Panel6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel6.Click
        Form1.Show()
    End Sub

    Private Sub Panel6_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel6.Paint

    End Sub

    Private Sub Panel8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel8.Click

    End Sub

    Private Sub Panel8_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel8.Paint

    End Sub

    Private Sub OvalShape1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OvalShape1.Click
        Form4.Show()
    End Sub

    Private Sub Panel5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel5.Click
        Close()
    End Sub

    Private Sub Panel5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel5.Paint

    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Close()

    End Sub
End Class