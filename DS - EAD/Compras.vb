Public Class Form1

    Private Sub AddcarButton_Click(sender As Object, e As EventArgs) Handles AddcarButton.Click

        If VerificaCampo(TextBox1.Text) = False Then

            MsgBox("Informe o Produto!", vbExclamation, "Aviso")

            Exit Sub

        End If

        Lb1.Items.Add(TextBox1.Text)
        Lb2.Items.Add(TextBox2.Text)
        Lb3.Items.Add(TextBox3.Text)
        Lb4.Items.Add(TextBox4.Text)

        If VerificaCampo(TextBox1.Text) = False Then
            MsgBox("Informe o Produto!", vbExclamation, "Aviso")

            Exit Sub

        End If

        If VerificaCampo(TextBox2.Text) = False Then
            MsgBox("Informe a Quantidade!", vbExclamation, "Aviso")

            Exit Sub

        End If

        If VerificaCampo(TextBox3.Text) = False Then
            MsgBox("Informe o Preço Unitário!", vbExclamation, "Aviso")

            Exit Sub

        End If

    End Sub

    Private Function VerificaCampo(Texto As String) As Boolean

        If Texto = "" Then
            VerificaCampo = False
        Else
            VerificaCampo = True
        End If


    End Function

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim val1, val2, res As Integer

        val1 = Val(TextBox2.Text)
        val2 = Val(TextBox3.Text)
        res = val1 * val2

        TextBox4.Text = res

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim del As Integer = Lb1.SelectedIndex
        If del <> -1 Then
            Lb1.Items.RemoveAt(del)
            Lb2.Items.RemoveAt(del)
            Lb3.Items.RemoveAt(del)
            Lb4.Items.RemoveAt(del)

        End If
    End Sub

    Private Sub Lb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lb1.SelectedIndexChanged
        Lb2.SelectedIndex = Lb1.SelectedIndex
        Lb3.SelectedIndex = Lb1.SelectedIndex
        Lb4.SelectedIndex = Lb1.SelectedIndex

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As Double
        Dim x As Integer

        For x = 0 To Lb4.Items.Count - 1
            result = result + CDbl(Lb4.Items(x))
        Next

        TextBox5.Text = result

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        Lb1.Items.Clear()
        Lb2.Items.Clear()
        Lb3.Items.Clear()
        Lb4.Items.Clear()
        TextBox5.Clear()

    End Sub

    Private Sub Lb2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lb2.SelectedIndexChanged
        Lb1.SelectedIndex = Lb2.SelectedIndex
        Lb3.SelectedIndex = Lb2.SelectedIndex
        Lb4.SelectedIndex = Lb2.SelectedIndex
    End Sub

    Private Sub Lb3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lb3.SelectedIndexChanged
        Lb1.SelectedIndex = Lb3.SelectedIndex
        Lb2.SelectedIndex = Lb3.SelectedIndex
        Lb4.SelectedIndex = Lb3.SelectedIndex
    End Sub

    Private Sub Lb4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lb4.SelectedIndexChanged
        Lb1.SelectedIndex = Lb4.SelectedIndex
        Lb2.SelectedIndex = Lb4.SelectedIndex
        Lb3.SelectedIndex = Lb4.SelectedIndex
    End Sub

End Class