Public Class FormLogin
    Private Sub ButtonLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLogin.Click
        If TextBoxUserId.Text = "" Then
            MsgBox("User ID Tidak Boleh Kosong")
            TextBoxUserId.Focus()
        ElseIf TextBoxPassword.Text = "" Then
            MsgBox("Password Tidak Boleh Kosong")
            TextBoxPassword.Focus()
        Else
            If UCase(TextBoxUserId.Text) = "ADMIN" And TextBoxPassword.Text = "123456" Then
                Me.Hide()
                MenuUtama.Show()
            Else
                MsgBox("USER ID/PASSWORD SALAH")
            End If
        End If
    End Sub

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        End
    End Sub

    Private Sub FormLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
