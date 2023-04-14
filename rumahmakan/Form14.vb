Imports System.Data.Odbc

Public Class Form14
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
        Dim sql As String = "insert into t_transaksi values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "','" & TextBox10.Text & "','" & TextBox11.Text & "','" & TextBox12.Text & "','" & TextBox13.Text & "','" & TextBox20.Text & "')"
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
            da = New OdbcDataAdapter("select *from t_transaksi order by id_transaksi asc",
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
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox1.Items.Add(dr.Item("id_costumer"))
        Loop
    End Sub
    Sub NOM()
        cmd = New OdbcCommand("select id_makanan from tb_makanan", con)
        dr = cmd.ExecuteReader
        ComboBox2.Items.Clear()
        Do While dr.Read
            ComboBox2.Items.Add(dr.Item("id_makanan"))
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
        Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

        End Sub

        Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

        End Sub

        Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

        End Sub

        Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
            cmd = New OdbcCommand("select * from tb_costumer where id_costumer='" & ComboBox1.Text & "'", con)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                TextBox2.Text = dr.Item("nm_costumer")
                TextBox3.Text = dr.Item("alm_costumer")
                TextBox4.Text = dr.Item("kota")
                TextBox5.Text = dr.Item("no_telepon")
            Else
                MsgBox("id minuman tidak ada")

            End If
        End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        cmd = New OdbcCommand("select * from tb_minuman where id_minuman='" & ComboBox3.Text & "'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox10.Text = dr.Item("nm_minuman")
            TextBox11.Text = dr.Item("harga_minuman")
            TextBox12.Text = dr.Item("nm_minuman2")
            TextBox13.Text = dr.Item("harga_minuman2")

        Else
            MsgBox("id minuman tidak ada")

        End If
    End Sub

    Private Sub Form14_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tampil()
        NIM()
        NOM()
        NEM()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        cmd = New OdbcCommand("select * from tb_makanan where id_makanan='" & ComboBox2.Text & "'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox6.Text = dr.Item("nm_makanan")
            TextBox7.Text = dr.Item("harga_makanan")
            TextBox8.Text = dr.Item("nm_makanan2")
            TextBox9.Text = dr.Item("harga_makanan2")

        Else
            MsgBox("id makanan tidak ada")

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        simpan()
    End Sub

    Private Sub TextBox15_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox15.TextChanged
        Dim hmk1 As Integer
        Dim hmk2 As Integer
        Dim mkb1 As Integer
        Dim mkb2 As Integer
        Dim tmak As Integer

        hmk1 = CInt(TextBox7.Text)
        hmk2 = CInt(TextBox9.Text)
        mkb1 = CInt(TextBox14.Text)
        mkb2 = CInt(TextBox15.Text)
        tmak = ((hmk1 * mkb1) + (hmk2 * mkb2))
        TextBox16.Text = tmak
    End Sub

    Private Sub TextBox18_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox18.TextChanged
        Dim hmn1 As Integer
        Dim hmn2 As Integer
        Dim mnb1 As Integer
        Dim mnb2 As Integer
        Dim tmin As Integer
       

        hmn1 = CInt(TextBox11.Text)
        hmn2 = CInt(TextBox13.Text)
        mnb1 = CInt(TextBox17.Text)
        mnb2 = CInt(TextBox18.Text)
        tmin = ((hmn1 * mnb1) + (hmn2 * mnb2))
       
        TextBox19.Text = tmin

    End Sub

    Private Sub TextBox19_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox19.TextChanged
        Dim minum As Integer
        Dim makan As Integer
        Dim total As Integer

        makan = CInt(TextBox16.Text)
        minum = CInt(TextBox19.Text)
        total = (makan + minum)

        TextBox20.Text = total
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        tampil()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim a As String = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        If a = "" Then
            MsgBox("Data pembayarn yang dihapus belum DIPILIH")
        Else
            If (MessageBox.Show("Anda yakin menghapus data dengan id_transaksi=" & a &
           "...?", "Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) =
           Windows.Forms.DialogResult.OK) Then
                koneksi()
                cmd = New OdbcCommand("delete from t_transaksi where id_transaksi='" & a &
               "'", con)
                cmd.ExecuteNonQuery()
                MsgBox("Menghapus data BERHASIL", vbInformation, "INFORMASI")
                con.Close()
                tampil()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = "0"
        TextBox8.Text = ""
        TextBox9.Text = "0"
        TextBox10.Text = ""
        TextBox11.Text = "0"
        TextBox12.Text = ""
        TextBox13.Text = "0"
        TextBox14.Text = "0"
        TextBox15.Text = "0"
        TextBox16.Text = "0"
        TextBox17.Text = "0"
        TextBox18.Text = "0"
        TextBox19.Text = "0"
        TextBox20.Text = "0"
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox1.Focus()
    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox20_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Close()
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click

    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click

    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click

    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click

    End Sub

    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click

    End Sub

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click

    End Sub

    Private Sub Label18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label18.Click

    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click

    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class