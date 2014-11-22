﻿
Imports System.Data.SQLite

Public Class frmResearchAgents

    Private ListColumnSorter As ListViewColumnSorter

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        lstAgents.Columns.Add("Agent", 130, HorizontalAlignment.Left)
        lstAgents.Columns.Add("Field", 150, HorizontalAlignment.Left)
        lstAgents.Columns.Add("Current RP", 80, HorizontalAlignment.Right)
        lstAgents.Columns.Add("Number of Cores", 95, HorizontalAlignment.Right)
        lstAgents.Columns.Add("Current Value", 90, HorizontalAlignment.Right)
        lstAgents.Columns.Add("RP/Day", 60, HorizontalAlignment.Right)
        lstAgents.Columns.Add("Level", 50, HorizontalAlignment.Center)
        lstAgents.Columns.Add("Location", 265, HorizontalAlignment.Left)

    End Sub

    Private Sub frmResearchAgents_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Call LoadGrid()
    End Sub

    ' Sort columns
    Private Sub lstAgents_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstAgents.ColumnClick
        ' Set the sort order options
        Call SetLstVwColumnSortOrder(e, ListColumnSorter)

        ' Perform the sort with these new sort options.
        lstAgents.Sort()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Hide()
    End Sub

    Private Sub LoadGrid()
        Dim lstViewRow As ListViewItem
        Dim fAccessError As New frmAPIError

        Dim readerPriceLookup As SQLiteDataReader
        Dim SQL As String
        Dim CurrentValue As Double
        Dim CurrentNumberofCores As Long
        Dim TotalValue As Double = 0

        lstAgents.Items.Clear()

        If SelectedCharacter.ResearchAccess Then

            Application.UseWaitCursor = True

            ListColumnSorter = New ListViewColumnSorter()
            lstAgents.ListViewItemSorter = ListColumnSorter

            lstAgents.BeginUpdate()

            With SelectedCharacter.GetResearchAgents
                For i = 0 To .GetResearchAgents.Count - 1
                    ' Get the total value of the datacores if I were to cash them in today - Price minus the DataCoreRedeemCost
                    SQL = "SELECT PRICE FROM ITEM_PRICES WHERE ITEM_NAME ='Datacore - " & .GetResearchAgents(i).Field & "'"
                    DBCommand = New SQLiteCommand(SQL, DB)
                    readerPriceLookup = DBCommand.ExecuteReader()

                    If readerPriceLookup.Read() Then
                        ' Get the number of cores we would get, minus the redeem cost from each 
                        CurrentNumberofCores = CLng(Math.Floor(.GetResearchAgents(i).CurrentRP / 100))
                        CurrentValue = Math.Floor(.GetResearchAgents(i).CurrentRP / 100) * (readerPriceLookup.GetDouble(0) - DataCoreRedeemCost)
                    Else
                        CurrentNumberofCores = 0
                        CurrentValue = 0
                    End If

                    ' Load the current data
                    lstViewRow = lstAgents.Items.Add(.GetResearchAgents(i).Agent) ' Agent Name
                    'The remaining columns are subitems  
                    lstViewRow.SubItems.Add(.GetResearchAgents(i).Field) ' Field
                    lstViewRow.SubItems.Add(FormatNumber(.GetResearchAgents(i).CurrentRP, 2)) ' Current RP
                    lstViewRow.SubItems.Add(FormatNumber(CurrentNumberofCores, 0)) ' Current number of cores
                    lstViewRow.SubItems.Add(FormatNumber(CurrentValue, 2)) ' Current Value
                    lstViewRow.SubItems.Add(FormatNumber(.GetResearchAgents(i).RPperDay, 2)) ' RP/Day
                    lstViewRow.SubItems.Add(CStr(.GetResearchAgents(i).AgentLevel)) ' Level
                    lstViewRow.SubItems.Add(.GetResearchAgents(i).Location) ' Location
                    TotalValue = TotalValue + CurrentValue
                Next
            End With

            ' Set total isk
            lblTotalDCValue.Text = FormatNumber(TotalValue, 2) & " ISK"

            ' Make sure the refresh button is enabled
            btnRefresh.Enabled = True
            lstAgents.EndUpdate()
            Application.UseWaitCursor = False
        Else
            Application.DoEvents()
            fAccessError.ErrorText = "This API did not allow research agents to be loaded for associated characters." & _
                Environment.NewLine & Environment.NewLine & "Please ensure your Customizable API includes 'Research' under the 'Science & Industry' section to include research agents and then reload the API."
            fAccessError.Text = "API: No Research Agents Loaded"
            fAccessError.ErrorLink = "http://support.eveonline.com/api/Key/CreatePredefined/589962/"
            fAccessError.ShowDialog()
            ' Disable the refresh button
            btnRefresh.Enabled = False
        End If
    End Sub

    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        Application.UseWaitCursor = True

        ' Reload the agents and update from API if necessary
        Call SelectedCharacter.GetResearchAgents.LoadResearchAgents(True)

        ' Refresh the data
        Call LoadGrid()

        Application.UseWaitCursor = False
        MsgBox("Records Refreshed", vbInformation, Application.ProductName)

    End Sub

End Class