Public Class Form12

    Private Sub Panel6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel6.Click
        Form1.Show()
    End Sub

    Private Sub Panel6_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel6.Paint

    End Sub

    Private Sub Panel8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel8.Click
        Form14.Show()
    End Sub

    Private Sub Panel8_ControlAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs) Handles Panel8.ControlAdded

    End Sub

    Private Sub Panel8_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel8.Paint

    End Sub

    Private Sub Panel7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel7.Click
        Form2.Show()
    End Sub

    Private Sub Panel7_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel7.Paint

    End Sub

    Private Sub OvalShape1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OvalShape1.Click
        Form13.Show()
    End Sub

    Private Sub Panel5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel5.Click
        Close()
        Form11.TextBox1.Text = ""
        Form11.TextBox2.Text = ""

    End Sub

    Private Sub Panel5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel5.Paint

    End Sub

    Private Sub Form12_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form15.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form18.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form8.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form16.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        form_report.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Form17.Show()
    End Sub
End Class