Public Class Form1
    Private Sub GenerateButton_Click(sender As Object, e As EventArgs) Handles GenerateButton.Click
        Dim passwordLength As Integer = Convert.ToInt32(LengthNumericUpDown.Value)
        Dim validChars As String = ""

        If UppercaseCheckBox.Checked Then
            validChars += "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        End If
        If LowercaseCheckBox.Checked Then
            validChars += "abcdefghijklmnopqrstuvwxyz"
        End If
        If NumbersCheckBox.Checked Then
            validChars += "0123456789"
        End If
        If SpecialCharsCheckBox.Checked Then
            validChars += "!@#$%^&*()_+"
        End If

        Dim password As String = GenerateRandomPassword(passwordLength, validChars)
        PasswordTextBox.Text = password
    End Sub

    Private Function GenerateRandomPassword(length As Integer, validChars As String) As String
        Dim random As New Random()
        Dim password As New System.Text.StringBuilder()

        For i As Integer = 1 To length
            Dim randomIndex As Integer = random.Next(0, validChars.Length)
            password.Append(validChars(randomIndex))
        Next

        Return password.ToString()
    End Function
End Class
