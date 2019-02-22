Public Class ErrorDialog

    Public ErrStackTrace As String

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        My.Application.Log.WriteEntry(Now().ToString & " " & Me.errMessage.Text.ToString)
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MsgBox(ErrStackTrace, MsgBoxStyle.Information, "Stack Trace")
    End Sub
End Class