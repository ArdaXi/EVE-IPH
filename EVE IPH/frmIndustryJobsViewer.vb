
Imports System.Data.SQLite

Public Class frmIndustryJobsViewer

    Private ColumnPositions(19) As String ' For saving the column order
    Private ColumnSorter As ListViewColumnSorter
    Private FirstLoad As Boolean
    Private AddingColumns As Boolean
    Private MovedColumn As Integer
    Private CurrentDateTime As Date

    Private Structure ColumnWidth
        Dim Name As String
        Dim Width As Integer
    End Structure

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        FirstLoad = True

        If UserApplicationSettings.ShowToolTips Then
            ttMain.SetToolTip(btnSaveSettings, "Saves Column order and Sort Column")
        End If

        CurrentDateTime = DateTime.UtcNow

        Timer1.Enabled = False
        Timer1.Interval = 1000 ' 1 second

        FirstLoad = False

    End Sub

    Private Sub frmIndustryJobsViewer_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Dim fAccessError As New frmAPIError

        FirstLoad = True

        ' See if they can load the jobs at all
        If Not SelectedCharacter.JobsAccess Then
            fAccessError.ErrorText = "This API did not allow industry jobs to be loaded for this character." & _
                Environment.NewLine & Environment.NewLine & "Please ensure your Customizable API includes 'IndustryJobs' under the 'Science & Industry' section to include industry jobs and then reload the API."
            fAccessError.Text = "API: No Industry Jobs Loaded"

            fAccessError.ErrorLink = "http://community.eveonline.com/support/api-key/CreatePredefined?accessMask=589962/"
            fAccessError.ShowDialog()

            gbInventionJobs.Enabled = False
        Else
            gbInventionJobs.Enabled = True

            If SelectedCharacter.CharacterCorporation.JobsAccess Then
                rbtnBothJobs.Enabled = True
                rbtnCorpJobs.Enabled = True
            Else
                rbtnBothJobs.Enabled = False
                rbtnCorpJobs.Enabled = False
            End If

            If UserIndustryJobsColumnSettings.ViewJobType = rbtnPersonalJobs.Text Then
                rbtnPersonalJobs.Checked = True
            ElseIf UserIndustryJobsColumnSettings.ViewJobType = rbtnCorpJobs.Text And rbtnCorpJobs.Enabled Then
                rbtnCorpJobs.Checked = True
            ElseIf UserIndustryJobsColumnSettings.ViewJobType = rbtnBothJobs.Text And rbtnBothJobs.Enabled Then
                rbtnBothJobs.Checked = True
            Else
                rbtnPersonalJobs.Checked = True
            End If

            If UserIndustryJobsColumnSettings.JobTimes = rbtnCurrentJobs.Text Then
                rbtnCurrentJobs.Checked = True
            Else
                rbtnJobHistory.Checked = True
            End If

            Call RefreshGrid()

            FirstLoad = False
        End If

    End Sub

    ' Refreshes main grid with the industry jobs
    Private Sub RefreshGrid()
        Dim SQL As String
        Dim rsJobs As SQLiteDataReader
        Dim lstJobRow As ListViewItem
        Dim JobState As String
        Dim JobStateColor As Color

        Dim StartDate As Date
        Dim EndDate As Date

        SQL = "SELECT activityName, status, startDate, endDate, completedDate, blueprintTypeName, "
        SQL = SQL & "CASE WHEN IT1.typeID <> 0 THEN IT1.typeName ELSE 'Unknown' END, "
        SQL = SQL & "CASE WHEN IT1.typeID <> 0 THEN INVENTORY_GROUPS.groupName ELSE 'Unknown' END, "
        SQL = SQL & "SOLAR_SYSTEMS.solarSystemName, regionName, licensedRuns, runs, successfulRuns, "
        SQL = SQL & "CASE WHEN S1.FACILITY_NAME IS NOT NULL THEN S1.FACILITY_NAME ELSE "
        SQL = SQL & "(CASE WHEN C1.FACILITY_NAME IS NOT NULL THEN C1.FACILITY_NAME || ' Container' ELSE "
        SQL = SQL & "(CASE WHEN IT2.typeName IS NOT NULL THEN IT2.typeName ELSE 'Unknown' END) END) END AS BPLID, "
        SQL = SQL & "CASE WHEN S2.FACILITY_NAME IS NOT NULL THEN S2.FACILITY_NAME ELSE "
        SQL = SQL & "(CASE WHEN C2.FACILITY_NAME IS NOT NULL THEN C2.FACILITY_NAME || ' Container' ELSE "
        SQL = SQL & "(CASE WHEN IT3.typeName IS NOT NULL THEN IT3.typeName ELSE 'Unknown' END) END) END AS OLID "
        SQL = SQL & "FROM INDUSTRY_JOBS "
        ' Stations
        SQL = SQL & "LEFT OUTER JOIN (SELECT FACILITY_ID, FACILITY_NAME FROM FACILITY_NAMES) AS S1 ON S1.FACILITY_ID = INDUSTRY_JOBS.blueprintLocationID "
        SQL = SQL & "LEFT OUTER JOIN (SELECT FACILITY_ID, FACILITY_NAME FROM FACILITY_NAMES) AS S2 ON S2.FACILITY_ID = INDUSTRY_JOBS.outputLocationID "
        ' Containers in stations
        SQL = SQL & "LEFT OUTER JOIN (SELECT FACILITY_ID, FACILITY_NAME, A1.ItemID FROM FACILITY_NAMES LEFT OUTER JOIN (SELECT LocationID, ItemID FROM ASSETS WHERE ID = " & SelectedCharacter.ID & ") "
        SQL = SQL & "AS A1 ON A1.LocationID = FACILITY_ID) AS C1 ON C1.ItemID = INDUSTRY_JOBS.blueprintLocationID "
        SQL = SQL & "LEFT OUTER JOIN (SELECT FACILITY_ID, FACILITY_NAME, A2.ItemID FROM FACILITY_NAMES LEFT OUTER JOIN (SELECT LocationID, ItemID FROM ASSETS WHERE ID = " & SelectedCharacter.ID & ") "
        SQL = SQL & "AS A2 ON A2.LocationID = FACILITY_ID) AS C2 ON C2.ItemID = INDUSTRY_JOBS.blueprintLocationID "
        ' POS modules
        SQL = SQL & "LEFT OUTER JOIN (SELECT typeID, typeName FROM INVENTORY_TYPES) AS IT2 ON IT2.typeID = INDUSTRY_JOBS.blueprintLocationID "
        SQL = SQL & "LEFT OUTER JOIN (SELECT typeID, typeName FROM INVENTORY_TYPES) AS IT3 ON IT3.typeID = INDUSTRY_JOBS.outputLocationID, "
        SQL = SQL & "RAM_ACTIVITIES, SOLAR_SYSTEMS, REGIONS, INVENTORY_TYPES AS IT1, INVENTORY_GROUPS "
        SQL = SQL & "WHERE INDUSTRY_JOBS.activityID = RAM_ACTIVITIES.activityID "
        SQL = SQL & "AND INDUSTRY_JOBS.solarSystemID = SOLAR_SYSTEMS.solarSystemID "
        SQL = SQL & "AND SOLAR_SYSTEMS.regionID = REGIONS.regionID "
        SQL = SQL & "AND INDUSTRY_JOBS.productTypeID = IT1.typeID "
        SQL = SQL & "AND IT1.groupID = INVENTORY_GROUPS.groupID "
        If rbtnCurrentJobs.Checked Then
            ' Only check status for current jobs
            SQL = SQL & "AND status <> 101 "
        End If
        SQL = SQL & "AND installerID = " & SelectedCharacter.ID & " "
        SQL = SQL & "AND JobType = 0 "

        If rbtnCorpJobs.Checked Then
            SQL = SQL & "AND JobType = " & CStr(ScanType.Corporation) & " "
        ElseIf rbtnPersonalJobs.Checked Then
            SQL = SQL & "AND JobType = " & CStr(ScanType.Personal) & " "
        End If

        ' Add sorting options here
        DBCommand = New SQLiteCommand(SQL, DB)
        rsJobs = DBCommand.ExecuteReader

        Application.UseWaitCursor = True
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        lstIndustryJobs.BeginUpdate()
        Call RefreshColumns()

        ' Create an instance of a ListView column sorter and assign it 
        ColumnSorter = New ListViewColumnSorter()
        lstIndustryJobs.ListViewItemSorter = ColumnSorter

        While rsJobs.Read

            StartDate = CDate(rsJobs.GetString(2))
            EndDate = CDate(rsJobs.GetString(3))

            ' Job State Flag
            If EndDate <= CurrentDateTime Then
                ' Job is done
                JobState = "Complete"
                JobStateColor = Color.Green
            ElseIf StartDate <= CurrentDateTime Then
                JobState = "In Progress"
                JobStateColor = Color.DarkOrange
            Else
                JobState = "Pending"
                JobStateColor = Color.Red
            End If

            If rsJobs.GetInt32(1) = 101 Then
                ' This has been completed
                JobState = "Completed"
                JobStateColor = Color.DarkGray
            End If

            ' Always add the end time to column 0
            lstJobRow = lstIndustryJobs.Items.Add(rsJobs.GetString(3))
            lstJobRow.UseItemStyleForSubItems = False

            With UserIndustryJobsColumnSettings

                For i = 1 To ColumnPositions.Count - 1

                    Select Case ColumnPositions(i)
                        Case ProgramSettings.JobStateColumn
                            lstJobRow.SubItems.Add(JobState) ' Job State
                            lstJobRow.SubItems(Array.IndexOf(ColumnPositions, "Job State")).ForeColor = JobStateColor
                        Case ProgramSettings.TimetoCompleteColumn
                            If JobState <> "Complete" And JobState <> "Completed" Then
                                lstJobRow.SubItems.Add(GetTimeToComplete(EndDate, CurrentDateTime))
                            Else
                                lstJobRow.SubItems.Add("")
                            End If
                        Case ProgramSettings.ActivityColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(0))
                        Case ProgramSettings.StatusColumn
                            lstJobRow.SubItems.Add(CStr(rsJobs.GetInt32(1)))
                        Case ProgramSettings.StartTimeColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(2))
                        Case ProgramSettings.EndTimeColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(3))
                        Case ProgramSettings.CompletionTimeColumn
                            If JobState = "Completed" Then
                                lstJobRow.SubItems.Add(rsJobs.GetString(4))
                            Else
                                lstJobRow.SubItems.Add("")
                            End If
                        Case ProgramSettings.BlueprintColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(5))
                        Case ProgramSettings.OutputItemColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(6))
                        Case ProgramSettings.OutputItemTypeColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(7))
                        Case ProgramSettings.InstallSolarSystemColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(8))
                        Case ProgramSettings.InstallRegionColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(9))
                        Case ProgramSettings.LicensedRunsColumn
                            lstJobRow.SubItems.Add(CStr(rsJobs.GetInt32(10)))
                        Case ProgramSettings.RunsColumn
                            lstJobRow.SubItems.Add(CStr(rsJobs.GetInt32(11)))
                        Case ProgramSettings.SuccessfulRunsColumn
                            lstJobRow.SubItems.Add(CStr(rsJobs.GetInt32(12)))
                        Case ProgramSettings.BlueprintLocationColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(13))
                        Case ProgramSettings.OutputLocationColumn
                            lstJobRow.SubItems.Add(rsJobs.GetString(14))

                    End Select
                Next
            End With
        End While

        lstIndustryJobs.EndUpdate()

        ' Force a click on the users column sort
        Dim e As New System.Windows.Forms.ColumnClickEventArgs(UserIndustryJobsColumnSettings.OrderByColumn)

        Select Case UserIndustryJobsColumnSettings.OrderType
            Case "Ascending"
                ColumnSorter.Order = SortOrder.Ascending
            Case "Descending"
                ColumnSorter.Order = SortOrder.Descending
            Case Else
                ColumnSorter.Order = SortOrder.None
        End Select

        ' Set the sort order options
        Call SetLstVwColumnSortOrder(e, ColumnSorter)

        ' Perform the sort with these new sort options.
        lstIndustryJobs.Sort()

        ' Enable the timer
        Timer1.Enabled = True

        Application.UseWaitCursor = False
        Me.Cursor = Cursors.Default
        Application.DoEvents()

    End Sub

    ' Clears the list and rebuilds it with columns they selected
    Private Sub RefreshColumns()

        Call LoadIndustryJobColumnPositions()

        Call lstIndustryJobs.Clear()
        AddingColumns = True

        ' Add an empty Column
        lstIndustryJobs.Columns.Add("")

        ' Now load all the columns in order of the settings
        For i = 1 To ColumnPositions.Count - 1
            If ColumnPositions(i) <> "" Then
                lstIndustryJobs.Columns.Add(ColumnPositions(i), GetColumnWidth(ColumnPositions(i)), GetColumnAlignment(ColumnPositions(i)))
            End If
        Next

        ' Empty Column not visible
        lstIndustryJobs.Columns(0).Width = 0

        AddingColumns = False

    End Sub

    ' Takes the column settings and saves the order to an array
    Private Sub LoadIndustryJobColumnPositions()

        For i = 0 To ColumnPositions.Count - 1
            ColumnPositions(i) = ""
        Next

        With UserIndustryJobsColumnSettings
            ColumnPositions(.JobState) = ProgramSettings.JobStateColumn
            ColumnPositions(.TimeToComplete) = ProgramSettings.TimetoCompleteColumn
            ColumnPositions(.Activity) = ProgramSettings.ActivityColumn
            ColumnPositions(.Status) = ProgramSettings.StatusColumn
            ColumnPositions(.StartTime) = ProgramSettings.StartTimeColumn
            ColumnPositions(.EndTime) = ProgramSettings.EndTimeColumn
            ColumnPositions(.CompletionTime) = ProgramSettings.CompletionTimeColumn
            ColumnPositions(.Blueprint) = ProgramSettings.BlueprintColumn
            ColumnPositions(.OutputItem) = ProgramSettings.OutputItemColumn
            ColumnPositions(.OutputItemType) = ProgramSettings.OutputItemTypeColumn
            ColumnPositions(.InstallSystem) = ProgramSettings.InstallSolarSystemColumn
            ColumnPositions(.InstallRegion) = ProgramSettings.InstallRegionColumn
            ColumnPositions(.LicensedRuns) = ProgramSettings.LicensedRunsColumn
            ColumnPositions(.Runs) = ProgramSettings.RunsColumn
            ColumnPositions(.SuccessfulRuns) = ProgramSettings.SuccessfulRunsColumn
            ColumnPositions(.BlueprintLocation) = ProgramSettings.BlueprintLocationColumn
            ColumnPositions(.OutputLocation) = ProgramSettings.OutputLocationColumn
        End With

        ' Reset the first one with nothing since the first column is empty
        ColumnPositions(0) = ""

    End Sub

    ' Returns the column Width for the sent column name
    Private Function GetColumnWidth(ColumnName As String) As Integer

        With UserIndustryJobsColumnSettings
            Select Case ColumnName
                Case ProgramSettings.JobStateColumn
                    Return .JobStateWidth
                Case ProgramSettings.TimetoCompleteColumn
                    Return .TimeToCompleteWidth
                Case ProgramSettings.ActivityColumn
                    Return .ActivityWidth
                Case ProgramSettings.StatusColumn
                    Return .StatusWidth
                Case ProgramSettings.StartTimeColumn
                    Return .StartTimeWidth
                Case ProgramSettings.EndTimeColumn
                    Return .EndTimeWidth
                Case ProgramSettings.CompletionTimeColumn
                    Return .CompletionTimeWidth
                Case ProgramSettings.BlueprintColumn
                    Return .BlueprintWidth
                Case ProgramSettings.OutputItemColumn
                    Return .OutputItemWidth
                Case ProgramSettings.OutputItemTypeColumn
                    Return .OutputItemTypeWidth
                Case ProgramSettings.InstallSolarSystemColumn
                    Return .InstallSystemWidth
                Case ProgramSettings.InstallRegionColumn
                    Return .InstallRegionWidth
                Case ProgramSettings.LicensedRunsColumn
                    Return .LicensedRunsWidth
                Case ProgramSettings.RunsColumn
                    Return .RunsWidth
                Case ProgramSettings.SuccessfulRunsColumn
                    Return .SuccessfulRunsWidth
                Case ProgramSettings.BlueprintLocationColumn
                    Return .BlueprintLocationWidth
                Case ProgramSettings.OutputLocationColumn
                    Return .OutputLocationWidth
                Case Else
                    Return 0
            End Select
        End With

    End Function

    ' Returns the allignment for the column name sent
    Private Function GetColumnAlignment(ColumnName As String) As System.Windows.Forms.HorizontalAlignment

        Select Case ColumnName
            Case ProgramSettings.JobStateColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.TimetoCompleteColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.ActivityColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.StartTimeColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.EndTimeColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.CompletionTimeColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.BlueprintColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.OutputItemColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.OutputItemTypeColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.InstallSolarSystemColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.InstallRegionColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.LicensedRunsColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.RunsColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.SuccessfulRunsColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.BlueprintLocationColumn
                Return HorizontalAlignment.Left
            Case ProgramSettings.OutputLocationColumn
                Return HorizontalAlignment.Left
            Case Else
                Return 0
        End Select

    End Function

    ' Updates the column order when changed
    Private Sub lstIndustryJobs_ColumnReordered(sender As Object, e As System.Windows.Forms.ColumnReorderedEventArgs) Handles lstIndustryJobs.ColumnReordered
        Dim TempArray(19) As String
        Dim Minus1 As Boolean = False

        e.Cancel = True ' Cancel the event so we can manually update the grid columns

        ' First index is ""
        TempArray(0) = ""

        If e.OldDisplayIndex > e.NewDisplayIndex Then
            ' For all indices larger than the new index, need to move it to the next array
            For i = 1 To e.NewDisplayIndex - 1
                TempArray(i) = ColumnPositions(i)
            Next

            ' Insert the new column
            TempArray(e.NewDisplayIndex) = ColumnPositions(e.OldDisplayIndex)

            ' Move all the rest of the items up one
            For i = e.NewDisplayIndex + 1 To TempArray.Count - 1
                If i < e.OldDisplayIndex + 1 Then
                    TempArray(i) = ColumnPositions(i - 1)
                Else
                    TempArray(i) = ColumnPositions(i)
                End If
            Next
        Else
            ' For all indices larger than the new index, need to move it to the next array
            For i = 1 To e.OldDisplayIndex - 1
                TempArray(i) = ColumnPositions(i)
            Next

            ' Insert the new column
            TempArray(e.NewDisplayIndex) = ColumnPositions(e.OldDisplayIndex)

            ' Back fill the array between the column we moved and the new location
            For i = e.OldDisplayIndex To e.NewDisplayIndex - 1
                TempArray(i) = ColumnPositions(i + 1)
            Next

            ' Replace all the items left
            For i = e.NewDisplayIndex + 1 To TempArray.Count - 1
                TempArray(i) = ColumnPositions(i)
            Next

        End If

        ColumnPositions = TempArray

        ' Savel the columns based on the current order
        With UserIndustryJobsColumnSettings
            For i = 1 To ColumnPositions.Count - 1
                Select Case ColumnPositions(i)
                    Case ProgramSettings.JobStateColumn
                        .JobState = i
                    Case ProgramSettings.TimetoCompleteColumn
                        .TimeToComplete = i
                    Case ProgramSettings.ActivityColumn
                        .Activity = i
                    Case ProgramSettings.StatusColumn
                        .Status = i
                    Case ProgramSettings.StartTimeColumn
                        .StartTime = i
                    Case ProgramSettings.EndTimeColumn
                        .EndTime = i
                    Case ProgramSettings.CompletionTimeColumn
                        .CompletionTime = i
                    Case ProgramSettings.BlueprintColumn
                        .Blueprint = i
                    Case ProgramSettings.OutputItemColumn
                        .OutputItem = i
                    Case ProgramSettings.OutputItemTypeColumn
                        .OutputItemType = i
                    Case ProgramSettings.InstallSolarSystemColumn
                        .InstallSystem = i
                    Case ProgramSettings.InstallRegionColumn
                        .InstallRegion = i
                    Case ProgramSettings.LicensedRunsColumn
                        .LicensedRuns = i
                    Case ProgramSettings.RunsColumn
                        .Runs = i
                    Case ProgramSettings.SuccessfulRunsColumn
                        .SuccessfulRuns = i
                    Case ProgramSettings.BlueprintLocationColumn
                        .BlueprintLocation = i
                    Case ProgramSettings.OutputLocationColumn
                        .OutputLocation = i
                End Select
            Next
        End With

        ' Now Refresh the grid
        Call RefreshGrid()

    End Sub

    ' Updates the column sizes when changed
    Private Sub lstIndustryJobs_ColumnWidthChanged(sender As Object, e As System.Windows.Forms.ColumnWidthChangedEventArgs) Handles lstIndustryJobs.ColumnWidthChanged
        Dim NewWidth As Integer = lstIndustryJobs.Columns(e.ColumnIndex).Width

        If Not AddingColumns Then
            With UserIndustryJobsColumnSettings
                Select Case ColumnPositions(e.ColumnIndex)
                    Case ProgramSettings.JobStateColumn
                        .JobStateWidth = NewWidth
                    Case ProgramSettings.TimetoCompleteColumn
                        .TimeToCompleteWidth = NewWidth
                    Case ProgramSettings.ActivityColumn
                        .ActivityWidth = NewWidth
                    Case ProgramSettings.StatusColumn
                        .StatusWidth = NewWidth
                    Case ProgramSettings.StartTimeColumn
                        .StartTimeWidth = NewWidth
                    Case ProgramSettings.EndTimeColumn
                        .EndTimeWidth = NewWidth
                    Case ProgramSettings.CompletionTimeColumn
                        .CompletionTimeWidth = NewWidth
                    Case ProgramSettings.BlueprintColumn
                        .BlueprintWidth = NewWidth
                    Case ProgramSettings.OutputItemColumn
                        .OutputItemWidth = NewWidth
                    Case ProgramSettings.OutputItemTypeColumn
                        .OutputItemTypeWidth = NewWidth
                    Case ProgramSettings.InstallSolarSystemColumn
                        .InstallSystemWidth = NewWidth
                    Case ProgramSettings.InstallRegionColumn
                        .InstallRegionWidth = NewWidth
                    Case ProgramSettings.LicensedRunsColumn
                        .LicensedRunsWidth = NewWidth
                    Case ProgramSettings.RunsColumn
                        .RunsWidth = NewWidth
                    Case ProgramSettings.SuccessfulRunsColumn
                        .SuccessfulRunsWidth = NewWidth
                    Case ProgramSettings.BlueprintLocationColumn
                        .BlueprintLocationWidth = NewWidth
                    Case ProgramSettings.OutputLocationColumn
                        .OutputLocationWidth = NewWidth
                End Select
            End With
        End If

    End Sub

    ' Determines if we display the sent colum
    Private Function ShowColumn(ColumnName As String) As Boolean
        If Array.IndexOf(ColumnPositions, ColumnName) <> -1 Then
            Return True
        Else
            Return False
        End If
    End Function

    ' Save the column order, the column size and the sort order
    Private Sub btnSaveSettings_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveSettings.Click
        If rbtnPersonalJobs.Checked Then
            UserIndustryJobsColumnSettings.ViewJobType = rbtnPersonalJobs.Text
        ElseIf rbtnCorpJobs.Checked Then
            UserIndustryJobsColumnSettings.ViewJobType = rbtnCorpJobs.Text
        ElseIf rbtnBothJobs.Checked Then
            UserIndustryJobsColumnSettings.ViewJobType = rbtnBothJobs.Text
        End If

        If rbtnCurrentJobs.Checked Then
            UserIndustryJobsColumnSettings.JobTimes = rbtnCurrentJobs.Text
        Else
            UserIndustryJobsColumnSettings.JobTimes = rbtnJobHistory.Text
        End If

        AllSettings.SaveIndustryJobsColumnSettings(UserIndustryJobsColumnSettings)

        MsgBox("Settings saved", vbInformation, Application.ProductName)

    End Sub

    Private Function GetTimeToComplete(EndJobDate As Date, CompareDate As Date) As String
        Dim SecondsDiff As Long

        SecondsDiff = DateDiff(DateInterval.Second, CompareDate, EndJobDate)

        Return FormatTimeToComplete(SecondsDiff)

    End Function

    Private Sub btnSelectColumns_Click(sender As System.Object, e As System.EventArgs) Handles btnSelectColumns.Click
        Dim f1 As New frmSelectIndustryJobColumns
        f1.ShowDialog()

        ' And refresh the Grid
        Call RefreshGrid()

    End Sub

    Private Sub btnUpdateJobs_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdateJobs.Click
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        ' Update jobs from API
        If rbtnPersonalJobs.Checked Or rbtnBothJobs.Checked Then
            ' Load the personal jobs
            Call SelectedCharacter.GetIndustryJobs.LoadIndustryJobs(ScanType.Personal, True)
        End If

        If rbtnBothJobs.Checked Or rbtnCorpJobs.Checked Then
            ' Load the corp jobs
            Call SelectedCharacter.CharacterCorporation.GetIndustryJobs.LoadIndustryJobs(ScanType.Corporation, True)
        End If

        Me.Cursor = Cursors.Default
        Application.DoEvents()

        MsgBox("Industry Jobs updated.", vbInformation, Application.ProductName)

    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Timer1.Enabled = False
        Me.Hide()
    End Sub

    Private Sub lstIndustryJobs_ColumnClick(sender As Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles lstIndustryJobs.ColumnClick
        ' Set the sort order options
        Call SetLstVwColumnSortOrder(e, ColumnSorter)

        ' Perform the sort with these new sort options.
        lstIndustryJobs.Sort()

        ' Save the column clicked
        UserIndustryJobsColumnSettings.OrderByColumn = e.Column

        Select Case ColumnSorter.Order
            Case SortOrder.Ascending
                UserIndustryJobsColumnSettings.OrderType = "Ascending"
            Case SortOrder.Descending
                UserIndustryJobsColumnSettings.OrderType = "Descending"
            Case Else
                UserIndustryJobsColumnSettings.OrderType = None
        End Select

    End Sub

    Private Sub btnRefreshList_Click(sender As System.Object, e As System.EventArgs) Handles btnRefreshList.Click
        Call RefreshGrid()
    End Sub

    Private Sub rbtnPersonalJobs_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnPersonalJobs.CheckedChanged
        If Not FirstLoad Then
            Call RefreshGrid()
        End If
    End Sub

    Private Sub rbtnCorpJobs_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnCorpJobs.CheckedChanged
        If Not FirstLoad Then
            Call RefreshGrid()
        End If
    End Sub

    Private Sub rbtnBothJobs_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnBothJobs.CheckedChanged
        If Not FirstLoad Then
            Call RefreshGrid()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Dim TimeToComplete As String
        Dim EndDate As Date

        ' On each tick just update the time column manually
        With UserIndustryJobsColumnSettings
            If .TimeToComplete <> 0 Then ' only if the time to complete column is visible
                CurrentDateTime = DateAdd(DateInterval.Second, 1, CurrentDateTime)

                For i = 0 To lstIndustryJobs.Items.Count - 1
                    ' Only update records with a time

                    If lstIndustryJobs.Items(i).SubItems(.JobState).Text <> "Complete" And lstIndustryJobs.Items(i).SubItems(.JobState).Text <> "Completed" Then

                        TimeToComplete = lstIndustryJobs.Items(i).SubItems(.TimeToComplete).Text
                        EndDate = CDate(lstIndustryJobs.Items(i).SubItems(0).Text)

                        lstIndustryJobs.Items(i).SubItems(.TimeToComplete).Text = GetTimeToComplete(EndDate, CurrentDateTime)
                        lstIndustryJobs.Update()
                        Application.DoEvents()
                    End If
                Next
            End If

        End With

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        Timer1.Enabled = False
    End Sub

    Private Sub rbtnCurrentJobs_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnCurrentJobs.CheckedChanged
        If Not FirstLoad And rbtnCurrentJobs.Checked Then
            Call RefreshGrid()
        End If
    End Sub

    Private Sub rbtnJobHistory_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnJobHistory.CheckedChanged
        If Not FirstLoad And rbtnJobHistory.Checked Then
            Call RefreshGrid()
        End If
    End Sub

End Class