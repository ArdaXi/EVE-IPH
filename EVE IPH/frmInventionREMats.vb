Public Class frmInventionREMats

    Public MaterialList As Materials
    Public InventionREUsage As Double
    Public MatType As String

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        If rbtnExportCSV.Text = UserApplicationSettings.DataExportFormat Then
            rbtnExportCSV.Checked = True
        ElseIf rbtnExportSSV.Text = UserApplicationSettings.DataExportFormat Then
            rbtnExportSSV.Checked = True
        ElseIf rbtnExportDefault.Text = UserApplicationSettings.DataExportFormat Then
            rbtnExportDefault.Checked = True
        End If

    End Sub

    Private Sub frmInventionREMats_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    Private Sub frmInventionREMats_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Dim MatList As ListViewItem
        Dim Quantity As Long
        Dim TotalCost As Double

        Application.UseWaitCursor = True
        Application.DoEvents()

        Me.Text = MatType

        ' Sort by quantity
        Call MaterialList.SortMaterialListByQuantity()

        ' Just load the mats into the list
        lstMats.Columns.Add("Material", 260, HorizontalAlignment.Left)
        lstMats.Columns.Add("Cost per Item", 100, HorizontalAlignment.Right)
        lstMats.Columns.Add("Total Cost", 100, HorizontalAlignment.Right)
        lstMats.Columns.Add("Quantity", 75, HorizontalAlignment.Right)

        For i = 0 To MaterialList.GetMaterialList.Count - 1
            Application.DoEvents()
            MatList = lstMats.Items.Add(MaterialList.GetMaterialList(i).GetMaterialName)
            TotalCost = CDbl(MaterialList.GetMaterialList(i).GetTotalCost)
            Quantity = CLng(MaterialList.GetMaterialList(i).GetQuantity)
            MatList.SubItems.Add(FormatNumber(TotalCost / Quantity, 2))
            MatList.SubItems.Add(FormatNumber(TotalCost, 2))
            MatList.SubItems.Add(FormatNumber(Quantity, 0))
        Next

        ' Finally add the usage cost
        MatList = lstMats.Items.Add("Total Usage")
        MatList.SubItems.Add(FormatNumber(InventionREUsage, 2))
        MatList.SubItems.Add(FormatNumber(InventionREUsage, 2))
        MatList.SubItems.Add(FormatNumber(1, 0))

        Application.UseWaitCursor = False

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Dispose()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub btnCopyMats_Click(sender As System.Object, e As System.EventArgs) Handles btnCopyMats.Click
        Dim TempExportType As String

        If rbtnExportCSV.Checked Then
            TempExportType = CSVDataExport
        ElseIf rbtnExportSSV.Checked Then
            TempExportType = SSVDataExport
        Else
            TempExportType = DefaultTextDataExport
        End If

        ' Paste to clipboard
        Call CopyTextToClipboard(MaterialList.GetClipboardList(TempExportType, True))
    End Sub

End Class