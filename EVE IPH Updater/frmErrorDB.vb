Public Class frmErrorDB

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Hide()
    End Sub

    Private Sub llblAccessDBLink_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblAccessDBLink.LinkClicked
        System.Diagnostics.Process.Start("https://sourceforge.net/projects/eveiph/")
    End Sub

    Private Sub frmErrorDB_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.Activate()
    End Sub
End Class