Public Class LoginForm
    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If TextBox1.Text <> "Administrator" Then
            ErrorProvider1.SetError(TextBox1, "Invalid Username")
            TextBox1.Text = ""
            TextBox1.Focus()
        End If
    End Sub
    Private Sub Blogin_Click(sender As Object, e As EventArgs) Handles Blogin.Click
        If TextBox2.Text <> "admin" Then
            ErrorProvider1.SetError(TextBox2, "Incorrect Password")
            TextBox2.Text = ""
            TextBox2.Focus()
        Else
            Form1.Show()
            TextBox1.Text = ""
            TextBox2.Text = ""
        End If
    End Sub
End Class