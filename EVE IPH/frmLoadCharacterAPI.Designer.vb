<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLoadCharacterAPI
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLoadCharacterAPI))
        Me.btnImportAPI = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.txtAPIKey = New System.Windows.Forms.TextBox()
        Me.lblAPIKey = New System.Windows.Forms.Label()
        Me.lblUserID = New System.Windows.Forms.Label()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.lblAPINote = New System.Windows.Forms.Label()
        Me.lblAPIAddress = New System.Windows.Forms.LinkLabel()
        Me.gbSelectChars = New System.Windows.Forms.GroupBox()
        Me.lblCorporationName = New System.Windows.Forms.Label()
        Me.chkCharacter3 = New System.Windows.Forms.CheckBox()
        Me.chkCharacter2 = New System.Windows.Forms.CheckBox()
        Me.chkCharacter1 = New System.Windows.Forms.CheckBox()
        Me.lblErrorText = New System.Windows.Forms.Label()
        Me.btnPrevious = New System.Windows.Forms.Button()
        Me.lblKeyType = New System.Windows.Forms.Label()
        Me.linklabelPredefined = New System.Windows.Forms.LinkLabel()
        Me.gbSelectChars.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnImportAPI
        '
        Me.btnImportAPI.Location = New System.Drawing.Point(220, 155)
        Me.btnImportAPI.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnImportAPI.Name = "btnImportAPI"
        Me.btnImportAPI.Size = New System.Drawing.Size(109, 26)
        Me.btnImportAPI.TabIndex = 0
        Me.btnImportAPI.Text = "Import"
        Me.btnImportAPI.UseVisualStyleBackColor = True
        Me.btnImportAPI.Visible = False
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(220, 155)
        Me.btnNext.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(109, 26)
        Me.btnNext.TabIndex = 4
        Me.btnNext.Text = "Next >"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(356, 155)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(109, 26)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'txtAPIKey
        '
        Me.txtAPIKey.Location = New System.Drawing.Point(84, 123)
        Me.txtAPIKey.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtAPIKey.Name = "txtAPIKey"
        Me.txtAPIKey.Size = New System.Drawing.Size(573, 22)
        Me.txtAPIKey.TabIndex = 3
        '
        'lblAPIKey
        '
        Me.lblAPIKey.AutoSize = True
        Me.lblAPIKey.Location = New System.Drawing.Point(16, 127)
        Me.lblAPIKey.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAPIKey.Name = "lblAPIKey"
        Me.lblAPIKey.Size = New System.Drawing.Size(61, 17)
        Me.lblAPIKey.TabIndex = 4
        Me.lblAPIKey.Text = "API Key:"
        '
        'lblUserID
        '
        Me.lblUserID.AutoSize = True
        Me.lblUserID.Location = New System.Drawing.Point(24, 92)
        Me.lblUserID.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblUserID.Name = "lblUserID"
        Me.lblUserID.Size = New System.Drawing.Size(53, 17)
        Me.lblUserID.TabIndex = 5
        Me.lblUserID.Text = "Key ID:"
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(84, 89)
        Me.txtUserID.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(108, 22)
        Me.txtUserID.TabIndex = 2
        '
        'lblAPINote
        '
        Me.lblAPINote.AutoSize = True
        Me.lblAPINote.Location = New System.Drawing.Point(19, 28)
        Me.lblAPINote.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAPINote.Name = "lblAPINote"
        Me.lblAPINote.Size = New System.Drawing.Size(200, 17)
        Me.lblAPINote.TabIndex = 7
        Me.lblAPINote.Text = "API Information available here:"
        '
        'lblAPIAddress
        '
        Me.lblAPIAddress.AutoSize = True
        Me.lblAPIAddress.Location = New System.Drawing.Point(216, 28)
        Me.lblAPIAddress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAPIAddress.Name = "lblAPIAddress"
        Me.lblAPIAddress.Size = New System.Drawing.Size(319, 17)
        Me.lblAPIAddress.TabIndex = 1
        Me.lblAPIAddress.TabStop = True
        Me.lblAPIAddress.Text = "https://community.eveonline.com/support/api-key/"
        '
        'gbSelectChars
        '
        Me.gbSelectChars.Controls.Add(Me.lblCorporationName)
        Me.gbSelectChars.Controls.Add(Me.chkCharacter3)
        Me.gbSelectChars.Controls.Add(Me.chkCharacter2)
        Me.gbSelectChars.Controls.Add(Me.chkCharacter1)
        Me.gbSelectChars.Controls.Add(Me.lblErrorText)
        Me.gbSelectChars.Location = New System.Drawing.Point(84, 36)
        Me.gbSelectChars.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gbSelectChars.Name = "gbSelectChars"
        Me.gbSelectChars.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gbSelectChars.Size = New System.Drawing.Size(436, 112)
        Me.gbSelectChars.TabIndex = 10
        Me.gbSelectChars.TabStop = False
        Me.gbSelectChars.Text = "Select Characters"
        Me.gbSelectChars.Visible = False
        '
        'lblCorporationName
        '
        Me.lblCorporationName.AutoSize = True
        Me.lblCorporationName.Location = New System.Drawing.Point(20, 30)
        Me.lblCorporationName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCorporationName.Name = "lblCorporationName"
        Me.lblCorporationName.Size = New System.Drawing.Size(79, 17)
        Me.lblCorporationName.TabIndex = 12
        Me.lblCorporationName.Text = "Corp Name"
        '
        'chkCharacter3
        '
        Me.chkCharacter3.Location = New System.Drawing.Point(24, 74)
        Me.chkCharacter3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkCharacter3.Name = "chkCharacter3"
        Me.chkCharacter3.Size = New System.Drawing.Size(384, 28)
        Me.chkCharacter3.TabIndex = 3
        Me.chkCharacter3.Text = "Character 3"
        Me.chkCharacter3.UseVisualStyleBackColor = True
        '
        'chkCharacter2
        '
        Me.chkCharacter2.Location = New System.Drawing.Point(24, 49)
        Me.chkCharacter2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkCharacter2.Name = "chkCharacter2"
        Me.chkCharacter2.Size = New System.Drawing.Size(384, 28)
        Me.chkCharacter2.TabIndex = 2
        Me.chkCharacter2.Text = "Character 2"
        Me.chkCharacter2.UseVisualStyleBackColor = True
        '
        'chkCharacter1
        '
        Me.chkCharacter1.Location = New System.Drawing.Point(24, 25)
        Me.chkCharacter1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkCharacter1.Name = "chkCharacter1"
        Me.chkCharacter1.Size = New System.Drawing.Size(384, 28)
        Me.chkCharacter1.TabIndex = 1
        Me.chkCharacter1.Text = "Character 1"
        Me.chkCharacter1.UseVisualStyleBackColor = True
        '
        'lblErrorText
        '
        Me.lblErrorText.Location = New System.Drawing.Point(19, 25)
        Me.lblErrorText.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblErrorText.Name = "lblErrorText"
        Me.lblErrorText.Size = New System.Drawing.Size(389, 73)
        Me.lblErrorText.TabIndex = 0
        Me.lblErrorText.Text = "Label1"
        '
        'btnPrevious
        '
        Me.btnPrevious.Enabled = False
        Me.btnPrevious.Location = New System.Drawing.Point(84, 155)
        Me.btnPrevious.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnPrevious.Name = "btnPrevious"
        Me.btnPrevious.Size = New System.Drawing.Size(109, 26)
        Me.btnPrevious.TabIndex = 6
        Me.btnPrevious.Text = "< Previous"
        Me.btnPrevious.UseVisualStyleBackColor = True
        '
        'lblKeyType
        '
        Me.lblKeyType.AutoSize = True
        Me.lblKeyType.Location = New System.Drawing.Point(19, 9)
        Me.lblKeyType.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblKeyType.Name = "lblKeyType"
        Me.lblKeyType.Size = New System.Drawing.Size(187, 17)
        Me.lblKeyType.TabIndex = 11
        Me.lblKeyType.Text = "Enter your Customizable API"
        '
        'linklabelPredefined
        '
        Me.linklabelPredefined.AutoSize = True
        Me.linklabelPredefined.Location = New System.Drawing.Point(216, 9)
        Me.linklabelPredefined.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.linklabelPredefined.Name = "linklabelPredefined"
        Me.linklabelPredefined.Size = New System.Drawing.Size(216, 17)
        Me.linklabelPredefined.TabIndex = 12
        Me.linklabelPredefined.TabStop = True
        Me.linklabelPredefined.Text = "(Click here for a pre-defined key)"
        '
        'frmLoadCharacterAPI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(672, 240)
        Me.ControlBox = False
        Me.Controls.Add(Me.linklabelPredefined)
        Me.Controls.Add(Me.lblKeyType)
        Me.Controls.Add(Me.btnPrevious)
        Me.Controls.Add(Me.lblAPIAddress)
        Me.Controls.Add(Me.lblAPINote)
        Me.Controls.Add(Me.txtUserID)
        Me.Controls.Add(Me.lblUserID)
        Me.Controls.Add(Me.lblAPIKey)
        Me.Controls.Add(Me.txtAPIKey)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.gbSelectChars)
        Me.Controls.Add(Me.btnImportAPI)
        Me.Controls.Add(Me.btnNext)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLoadCharacterAPI"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Import EVE API"
        Me.gbSelectChars.ResumeLayout(False)
        Me.gbSelectChars.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnImportAPI As System.Windows.Forms.Button
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents txtAPIKey As System.Windows.Forms.TextBox
    Friend WithEvents lblAPIKey As System.Windows.Forms.Label
    Friend WithEvents lblUserID As System.Windows.Forms.Label
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents lblAPINote As System.Windows.Forms.Label
    Friend WithEvents lblAPIAddress As System.Windows.Forms.LinkLabel
    Friend WithEvents gbSelectChars As System.Windows.Forms.GroupBox
    Friend WithEvents btnPrevious As System.Windows.Forms.Button
    Friend WithEvents lblErrorText As System.Windows.Forms.Label
    Friend WithEvents chkCharacter3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkCharacter2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkCharacter1 As System.Windows.Forms.CheckBox
    Friend WithEvents lblKeyType As System.Windows.Forms.Label
    Friend WithEvents lblCorporationName As System.Windows.Forms.Label
    Friend WithEvents linklabelPredefined As System.Windows.Forms.LinkLabel
End Class
