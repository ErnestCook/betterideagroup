Public Class frmStatus
    Dim lCancelrequested As Boolean = False

    Private Sub frmStatus_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub frmstatus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.MyProgressBar.Value = 0
        Me.MyProgressBar.Maximum = CInt(Me.lblMaximumRecord.Text)
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        lCancelrequested = True
    End Sub

    Public Property CancelRequested() As Boolean
        Get
            Return lCancelrequested
        End Get
        Set(ByVal value As Boolean)
            lCancelrequested = value
        End Set
    End Property

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class