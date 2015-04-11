<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCRESTMarketViewer
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
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series2 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series3 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series4 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series5 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Title1 As System.Windows.Forms.DataVisualization.Charting.Title = New System.Windows.Forms.DataVisualization.Charting.Title()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCRESTMarketViewer))
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.MarketChart = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.gbOptions = New System.Windows.Forms.GroupBox()
        Me.chkVolume = New System.Windows.Forms.CheckBox()
        Me.chkOrders = New System.Windows.Forms.CheckBox()
        Me.chkAvgPrice = New System.Windows.Forms.CheckBox()
        Me.chkHighPrice = New System.Windows.Forms.CheckBox()
        Me.chkLowPrice = New System.Windows.Forms.CheckBox()
        Me.DoubleTrackBar1 = New EVE_Isk_per_Hour.DoubleTrackBar()
        CType(Me.MarketChart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnImport
        '
        Me.btnImport.Enabled = False
        Me.btnImport.Location = New System.Drawing.Point(305, 470)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(84, 27)
        Me.btnImport.TabIndex = 0
        Me.btnImport.Text = "Import"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(485, 470)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(84, 27)
        Me.btnExit.TabIndex = 2
        Me.btnExit.Text = "Close"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'MarketChart
        '
        Me.MarketChart.BorderlineColor = System.Drawing.Color.Black
        ChartArea1.Name = "ChartArea1"
        Me.MarketChart.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.MarketChart.Legends.Add(Legend1)
        Me.MarketChart.Location = New System.Drawing.Point(12, 12)
        Me.MarketChart.Name = "MarketChart"
        Me.MarketChart.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.None
        Me.MarketChart.RightToLeft = System.Windows.Forms.RightToLeft.No
        Series1.ChartArea = "ChartArea1"
        Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series1.CustomProperties = "PointWidth=0.5"
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Series2.ChartArea = "ChartArea1"
        Series2.CustomProperties = "PointWidth=0.5"
        Series2.Legend = "Legend1"
        Series2.Name = "Series2"
        Series3.ChartArea = "ChartArea1"
        Series3.CustomProperties = "PointWidth=0.5"
        Series3.Legend = "Legend1"
        Series3.Name = "Series3"
        Series4.ChartArea = "ChartArea1"
        Series4.CustomProperties = "PointWidth=0.5"
        Series4.Legend = "Legend1"
        Series4.Name = "Series4"
        Series5.ChartArea = "ChartArea1"
        Series5.CustomProperties = "PointWidth=0.5"
        Series5.Legend = "Legend1"
        Series5.Name = "Series5"
        Me.MarketChart.Series.Add(Series1)
        Me.MarketChart.Series.Add(Series2)
        Me.MarketChart.Series.Add(Series3)
        Me.MarketChart.Series.Add(Series4)
        Me.MarketChart.Series.Add(Series5)
        Me.MarketChart.Size = New System.Drawing.Size(853, 428)
        Me.MarketChart.TabIndex = 3
        Me.MarketChart.Text = "Item Name"
        Title1.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Bold)
        Title1.IsDockedInsideChartArea = False
        Title1.Name = "Main Title"
        Title1.Text = "Item Name"
        Me.MarketChart.Titles.Add(Title1)
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(395, 470)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 27)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Save Settings"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'gbOptions
        '
        Me.gbOptions.Controls.Add(Me.chkVolume)
        Me.gbOptions.Controls.Add(Me.chkOrders)
        Me.gbOptions.Controls.Add(Me.chkAvgPrice)
        Me.gbOptions.Controls.Add(Me.chkHighPrice)
        Me.gbOptions.Controls.Add(Me.chkLowPrice)
        Me.gbOptions.Location = New System.Drawing.Point(591, 446)
        Me.gbOptions.Name = "gbOptions"
        Me.gbOptions.Size = New System.Drawing.Size(194, 65)
        Me.gbOptions.TabIndex = 5
        Me.gbOptions.TabStop = False
        '
        'chkVolume
        '
        Me.chkVolume.AutoSize = True
        Me.chkVolume.Location = New System.Drawing.Point(122, 14)
        Me.chkVolume.Name = "chkVolume"
        Me.chkVolume.Size = New System.Drawing.Size(61, 17)
        Me.chkVolume.TabIndex = 3
        Me.chkVolume.Text = "Volume"
        Me.chkVolume.UseVisualStyleBackColor = True
        '
        'chkOrders
        '
        Me.chkOrders.AutoSize = True
        Me.chkOrders.Location = New System.Drawing.Point(122, 30)
        Me.chkOrders.Name = "chkOrders"
        Me.chkOrders.Size = New System.Drawing.Size(57, 17)
        Me.chkOrders.TabIndex = 4
        Me.chkOrders.Text = "Orders"
        Me.chkOrders.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkOrders.UseVisualStyleBackColor = True
        '
        'chkAvgPrice
        '
        Me.chkAvgPrice.AutoSize = True
        Me.chkAvgPrice.Location = New System.Drawing.Point(14, 46)
        Me.chkAvgPrice.Name = "chkAvgPrice"
        Me.chkAvgPrice.Size = New System.Drawing.Size(93, 17)
        Me.chkAvgPrice.TabIndex = 2
        Me.chkAvgPrice.Text = "Average Price"
        Me.chkAvgPrice.UseVisualStyleBackColor = True
        '
        'chkHighPrice
        '
        Me.chkHighPrice.AutoSize = True
        Me.chkHighPrice.Location = New System.Drawing.Point(14, 30)
        Me.chkHighPrice.Name = "chkHighPrice"
        Me.chkHighPrice.Size = New System.Drawing.Size(75, 17)
        Me.chkHighPrice.TabIndex = 1
        Me.chkHighPrice.Text = "High Price"
        Me.chkHighPrice.UseVisualStyleBackColor = True
        '
        'chkLowPrice
        '
        Me.chkLowPrice.AutoSize = True
        Me.chkLowPrice.Location = New System.Drawing.Point(14, 14)
        Me.chkLowPrice.Name = "chkLowPrice"
        Me.chkLowPrice.Size = New System.Drawing.Size(73, 17)
        Me.chkLowPrice.TabIndex = 0
        Me.chkLowPrice.Text = "Low Price"
        Me.chkLowPrice.UseVisualStyleBackColor = True
        '
        'DoubleTrackBar1
        '
        Me.DoubleTrackBar1.Location = New System.Drawing.Point(12, 453)
        Me.DoubleTrackBar1.Maximum = 10
        Me.DoubleTrackBar1.Minimum = 0
        Me.DoubleTrackBar1.Name = "DoubleTrackBar1"
        Me.DoubleTrackBar1.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.DoubleTrackBar1.Size = New System.Drawing.Size(287, 57)
        Me.DoubleTrackBar1.SmallChange = 1
        Me.DoubleTrackBar1.TabIndex = 6
        Me.DoubleTrackBar1.Text = "DoubleTrackBar1"
        Me.DoubleTrackBar1.ValueLeft = 0
        Me.DoubleTrackBar1.ValueRight = 7
        '
        'frmCRESTMarketViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(877, 516)
        Me.Controls.Add(Me.DoubleTrackBar1)
        Me.Controls.Add(Me.gbOptions)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.MarketChart)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnImport)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCRESTMarketViewer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Market Viewer"
        CType(Me.MarketChart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbOptions.ResumeLayout(False)
        Me.gbOptions.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents MarketChart As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents gbOptions As System.Windows.Forms.GroupBox
    Friend WithEvents chkVolume As System.Windows.Forms.CheckBox
    Friend WithEvents chkOrders As System.Windows.Forms.CheckBox
    Friend WithEvents chkAvgPrice As System.Windows.Forms.CheckBox
    Friend WithEvents chkHighPrice As System.Windows.Forms.CheckBox
    Friend WithEvents chkLowPrice As System.Windows.Forms.CheckBox
    Friend WithEvents DoubleTrackBar1 As EVE_Isk_per_Hour.DoubleTrackBar
End Class
