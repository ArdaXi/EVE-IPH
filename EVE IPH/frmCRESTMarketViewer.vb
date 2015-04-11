

Public Class frmCRESTMarketViewer

    Public Sub New(WindowType As CopyPasteWindowType)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.


    End Sub

    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Me.Dispose()
    End Sub

End Class