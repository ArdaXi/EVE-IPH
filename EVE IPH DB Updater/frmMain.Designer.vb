<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnImages = New System.Windows.Forms.Button()
        Me.btnDatabase = New System.Windows.Forms.Button()
        Me.pgMain = New System.Windows.Forms.ProgressBar()
        Me.lblTableName = New System.Windows.Forms.Label()
        Me.btnBuildSQLServerDB = New System.Windows.Forms.Button()
        Me.TreeView = New System.Windows.Forms.TreeView()
        Me.SuspendLayout()
        '
        'btnImages
        '
        Me.btnImages.Location = New System.Drawing.Point(142, 15)
        Me.btnImages.Name = "btnImages"
        Me.btnImages.Size = New System.Drawing.Size(104, 29)
        Me.btnImages.TabIndex = 0
        Me.btnImages.Text = "Image Copy"
        Me.btnImages.UseVisualStyleBackColor = True
        '
        'btnDatabase
        '
        Me.btnDatabase.Location = New System.Drawing.Point(19, 15)
        Me.btnDatabase.Name = "btnDatabase"
        Me.btnDatabase.Size = New System.Drawing.Size(104, 29)
        Me.btnDatabase.TabIndex = 1
        Me.btnDatabase.Text = "Build DB"
        Me.btnDatabase.UseVisualStyleBackColor = True
        '
        'pgMain
        '
        Me.pgMain.Location = New System.Drawing.Point(19, 114)
        Me.pgMain.Name = "pgMain"
        Me.pgMain.Size = New System.Drawing.Size(226, 18)
        Me.pgMain.TabIndex = 2
        Me.pgMain.Visible = False
        '
        'lblTableName
        '
        Me.lblTableName.Location = New System.Drawing.Point(1, 93)
        Me.lblTableName.Name = "lblTableName"
        Me.lblTableName.Size = New System.Drawing.Size(253, 18)
        Me.lblTableName.TabIndex = 3
        Me.lblTableName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnBuildSQLServerDB
        '
        Me.btnBuildSQLServerDB.Location = New System.Drawing.Point(66, 52)
        Me.btnBuildSQLServerDB.Name = "btnBuildSQLServerDB"
        Me.btnBuildSQLServerDB.Size = New System.Drawing.Size(132, 29)
        Me.btnBuildSQLServerDB.TabIndex = 4
        Me.btnBuildSQLServerDB.Text = "Build SQL Server DB"
        Me.btnBuildSQLServerDB.UseVisualStyleBackColor = True
        '
        'TreeView
        '
        Me.TreeView.Location = New System.Drawing.Point(10, 87)
        Me.TreeView.Name = "TreeView"
        Me.TreeView.Size = New System.Drawing.Size(244, 24)
        Me.TreeView.TabIndex = 6
        Me.TreeView.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(264, 137)
        Me.Controls.Add(Me.TreeView)
        Me.Controls.Add(Me.btnBuildSQLServerDB)
        Me.Controls.Add(Me.lblTableName)
        Me.Controls.Add(Me.pgMain)
        Me.Controls.Add(Me.btnDatabase)
        Me.Controls.Add(Me.btnImages)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EVE Data Dump Updater"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnImages As System.Windows.Forms.Button
    Friend WithEvents btnDatabase As System.Windows.Forms.Button
    Friend WithEvents pgMain As System.Windows.Forms.ProgressBar
    Friend WithEvents lblTableName As System.Windows.Forms.Label
    Friend WithEvents btnBuildSQLServerDB As System.Windows.Forms.Button
    Friend WithEvents TreeView As System.Windows.Forms.TreeView

End Class
