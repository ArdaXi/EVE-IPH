Public Class frmSelectManufacturingTabColumns

    Dim MaxColumnNumber As Integer
    Dim SelectedIndex As Integer

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        MaxColumnNumber = 1
        SelectedIndex = 0

    End Sub

    Private Sub SetMaxColumnNumber(InNumber As Integer)
        If InNumber > MaxColumnNumber Then
            MaxColumnNumber = InNumber
        End If
    End Sub

    ' Load all the current checks
    Private Sub frmSelectIndustryJobColumns_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Call ShowList()
    End Sub

    Private Sub UpdateListCheck(ByVal ColumnValue As Integer, Index As Integer)
        If ColumnValue <> 0 Then
            chkLstBoxColumns.SetItemChecked(Index, True)
            SetMaxColumnNumber(ColumnValue)
        Else
            chkLstBoxColumns.SetItemChecked(Index, False)
        End If
    End Sub

    Private Sub ShowList()
        With UserManufacturingTabColumnSettings
            Call UpdateListCheck(.ItemCategory, 0)
            Call UpdateListCheck(.ItemGroup, 1)
            Call UpdateListCheck(.ItemName, 2)
            Call UpdateListCheck(.Owned, 3)
            Call UpdateListCheck(.Tech, 4)
            Call UpdateListCheck(.BPME, 5)
            Call UpdateListCheck(.BPTE, 6)
            Call UpdateListCheck(.Inputs, 7)
            Call UpdateListCheck(.Compared, 8)
            Call UpdateListCheck(.Runs, 9)
            Call UpdateListCheck(.ProductionLines, 10)
            Call UpdateListCheck(.LaboratoryLines, 11)
            Call UpdateListCheck(.InventionCost, 12)
            Call UpdateListCheck(.CopyCost, 13)
            Call UpdateListCheck(.TotalManufacturingCost, 14)
            Call UpdateListCheck(.Taxes, 15)
            Call UpdateListCheck(.BrokerFees, 16)
            Call UpdateListCheck(.BPProductionTime, 17)
            Call UpdateListCheck(.TotalProductionTime, 18)
            Call UpdateListCheck(.ItemMarketPrice, 19)
            Call UpdateListCheck(.Profit, 20)
            Call UpdateListCheck(.ProfitPercentage, 21)
            Call UpdateListCheck(.IskperHour, 22)
            Call UpdateListCheck(.SVR, 23)
            Call UpdateListCheck(.TotalCost, 24)
            Call UpdateListCheck(.BaseJobCost, 25)
            Call UpdateListCheck(.ManufacturingJobFee, 26)
            Call UpdateListCheck(.ManufacturingFacilityName, 27)
            Call UpdateListCheck(.ManufacturingFacilitySystem, 28)
            Call UpdateListCheck(.ManufacturingFacilitySystemIndex, 29)
            Call UpdateListCheck(.ManufacturingFacilityTax, 30)
            Call UpdateListCheck(.ManufacturingFacilityRegion, 31)
            Call UpdateListCheck(.ManufacturingFacilityMEBonus, 32)
            Call UpdateListCheck(.ManufacturingFacilityTEBonus, 33)
            Call UpdateListCheck(.ManufacturingFacilityUsage, 34)
            Call UpdateListCheck(.ComponentFacilityName, 35)
            Call UpdateListCheck(.ComponentFacilitySystem, 36)
            Call UpdateListCheck(.ComponentFacilitySystemIndex, 37)
            Call UpdateListCheck(.ComponentFacilityTax, 38)
            Call UpdateListCheck(.ComponentFacilityRegion, 39)
            Call UpdateListCheck(.ComponentFacilityMEBonus, 40)
            Call UpdateListCheck(.ComponentFacilityTEBonus, 41)
            Call UpdateListCheck(.ComponentFacilityUsage, 42)
            Call UpdateListCheck(.CopyingFacilityName, 43)
            Call UpdateListCheck(.CopyingFacilitySystem, 44)
            Call UpdateListCheck(.CopyingFacilitySystemIndex, 45)
            Call UpdateListCheck(.CopyingFacilityTax, 46)
            Call UpdateListCheck(.CopyingFacilityRegion, 47)
            Call UpdateListCheck(.CopyingFacilityMEBonus, 48)
            Call UpdateListCheck(.CopyingFacilityTEBonus, 49)
            Call UpdateListCheck(.CopyingFacilityUsage, 50)
            Call UpdateListCheck(.InventionFacilityName, 51)
            Call UpdateListCheck(.InventionFacilitySystem, 52)
            Call UpdateListCheck(.InventionFacilitySystemIndex, 53)
            Call UpdateListCheck(.InventionFacilityTax, 54)
            Call UpdateListCheck(.InventionFacilityRegion, 55)
            Call UpdateListCheck(.InventionFacilityMEBonus, 56)
            Call UpdateListCheck(.InventionFacilityTEBonus, 57)
            Call UpdateListCheck(.InventionFacilityUsage, 58)
            Call UpdateListCheck(.ManufacturingTeamName, 59)
            Call UpdateListCheck(.ManufacturingTeamBonuses, 60)
            Call UpdateListCheck(.ManufacturingTeamUsage, 61)
            Call UpdateListCheck(.ManufacturingTeamCostModifier, 62)
            Call UpdateListCheck(.ComponentTeamName, 63)
            Call UpdateListCheck(.ComponentTeamBonuses, 64)
            Call UpdateListCheck(.ComponentTeamUsage, 65)
            Call UpdateListCheck(.ComponentTeamCostModifier, 66)
            Call UpdateListCheck(.CopyingTeamName, 67)
            Call UpdateListCheck(.CopyingTeamBonuses, 68)
            Call UpdateListCheck(.CopyingTeamUsage, 69)
            Call UpdateListCheck(.CopyingTeamCostModifier, 70)
            Call UpdateListCheck(.InventionTeamName, 71)
            Call UpdateListCheck(.InventionTeamBonuses, 72)
            Call UpdateListCheck(.InventionTeamUsage, 73)
            Call UpdateListCheck(.InventionTeamCostModifier, 74)

            chkLstBoxColumns.Update()

        End With
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Hide()
    End Sub

    Private Sub btnSaveSettings_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveSettings.Click

        If chkLstBoxColumns.CheckedItems.Count = 0 Then
            MsgBox("You must select at least one Column", vbExclamation, Application.ProductName)
            Exit Sub
        End If

        ' Save the local settings and the user settings
        Call SaveLocalColumnSettings()

        ' Save the data in the XML file
        Call AllSettings.SaveManufacturingTabColumnSettings(UserManufacturingTabColumnSettings)

        MsgBox("Columns Saved", vbInformation, Application.ProductName)
        ManufacturingTabColumnsChanged = True

        Me.Hide()

    End Sub

    ' Processes the column order
    Private Function GetColumnNumber(ByVal ChkState As CheckState, CurrentValue As Integer) As Integer
        Dim NewValue As Integer

        ' Change to max column order + 1 if checked and not already set
        If CurrentValue = 0 And ChkState = CheckState.Checked Then
            NewValue = MaxColumnNumber + 1
            MaxColumnNumber += 1
        ElseIf ChkState = CheckState.Unchecked Then
            NewValue = 0
        Else
            NewValue = CurrentValue
        End If

        Return NewValue

    End Function

    ' Save the items as viewed or not and order them from the last column
    Public Sub SaveLocalColumnSettings()
        Dim ColumnPositions(75) As String
        Dim TempColumns(75) As String
        Dim ColumnCount As Integer = 0
        Dim i As Integer = 1
        Dim j As Integer = 1

        With UserManufacturingTabColumnSettings
            ' First add any new check boxes that weren't checked before
            .ItemCategory = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(0), .ItemCategory)
            .ItemGroup = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(1), .ItemGroup)
            .ItemName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(2), .ItemName)
            .Owned = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(3), .Owned)
            .Tech = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(4), .Tech)
            .BPME = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(5), .BPME)
            .BPTE = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(6), .BPTE)
            .Inputs = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(7), .Inputs)
            .Compared = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(8), .Compared)
            .Runs = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(9), .Runs)
            .ProductionLines = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(10), .ProductionLines)
            .LaboratoryLines = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(11), .LaboratoryLines)
            .InventionCost = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(12), .InventionCost)
            .CopyCost = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(13), .CopyCost)
            .TotalManufacturingCost = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(14), .TotalManufacturingCost)
            .Taxes = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(15), .Taxes)
            .BrokerFees = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(16), .BrokerFees)
            .BPProductionTime = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(17), .BPProductionTime)
            .TotalProductionTime = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(18), .TotalProductionTime)
            .ItemMarketPrice = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(19), .ItemMarketPrice)
            .Profit = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(20), .Profit)
            .ProfitPercentage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(21), .ProfitPercentage)
            .IskperHour = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(22), .IskperHour)
            .SVR = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(23), .SVR)
            .TotalCost = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(24), .TotalCost)
            .BaseJobCost = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(25), .BaseJobCost)
            .ManufacturingJobFee = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(26), .ManufacturingJobFee)
            .ManufacturingFacilityName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(27), .ManufacturingFacilityName)
            .ManufacturingFacilitySystem = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(28), .ManufacturingFacilitySystem)
            .ManufacturingFacilitySystemIndex = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(29), .ManufacturingFacilitySystemIndex)
            .ManufacturingFacilityTax = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(30), .ManufacturingFacilityTax)
            .ManufacturingFacilityRegion = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(31), .ManufacturingFacilityRegion)
            .ManufacturingFacilityMEBonus = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(32), .ManufacturingFacilityMEBonus)
            .ManufacturingFacilityTEBonus = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(33), .ManufacturingFacilityTEBonus)
            .ManufacturingFacilityUsage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(34), .ManufacturingFacilityUsage)
            .ComponentFacilityName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(35), .ComponentFacilityName)
            .ComponentFacilitySystem = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(36), .ComponentFacilitySystem)
            .ComponentFacilitySystemIndex = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(37), .ComponentFacilitySystemIndex)
            .ComponentFacilityTax = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(38), .ComponentFacilityTax)
            .ComponentFacilityRegion = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(39), .ComponentFacilityRegion)
            .ComponentFacilityMEBonus = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(40), .ComponentFacilityMEBonus)
            .ComponentFacilityTEBonus = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(41), .ComponentFacilityTEBonus)
            .ComponentFacilityUsage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(42), .ComponentFacilityUsage)
            .CopyingFacilityName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(43), .CopyingFacilityName)
            .CopyingFacilitySystem = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(44), .CopyingFacilitySystem)
            .CopyingFacilitySystemIndex = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(45), .CopyingFacilitySystemIndex)
            .CopyingFacilityTax = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(46), .CopyingFacilityTax)
            .CopyingFacilityRegion = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(47), .CopyingFacilityRegion)
            .CopyingFacilityMEBonus = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(48), .CopyingFacilityMEBonus)
            .CopyingFacilityTEBonus = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(49), .CopyingFacilityTEBonus)
            .CopyingFacilityUsage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(50), .CopyingFacilityUsage)
            .InventionFacilityName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(51), .InventionFacilityName)
            .InventionFacilitySystem = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(52), .InventionFacilitySystem)
            .InventionFacilitySystemIndex = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(53), .InventionFacilitySystemIndex)
            .InventionFacilityTax = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(54), .InventionFacilityTax)
            .InventionFacilityRegion = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(55), .InventionFacilityRegion)
            .InventionFacilityMEBonus = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(56), .InventionFacilityMEBonus)
            .InventionFacilityTEBonus = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(57), .InventionFacilityTEBonus)
            .InventionFacilityUsage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(58), .InventionFacilityUsage)
            .ManufacturingTeamName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(59), .ManufacturingTeamName)
            .ManufacturingTeamBonuses = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(60), .ManufacturingTeamBonuses)
            .ManufacturingTeamUsage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(61), .ManufacturingTeamUsage)
            .ManufacturingTeamCostModifier = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(62), .ManufacturingTeamCostModifier)
            .ComponentTeamName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(63), .ComponentTeamName)
            .ComponentTeamBonuses = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(64), .ComponentTeamBonuses)
            .ComponentTeamUsage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(65), .ComponentTeamUsage)
            .ComponentTeamCostModifier = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(66), .ComponentTeamCostModifier)
            .CopyingTeamName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(67), .CopyingTeamName)
            .CopyingTeamBonuses = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(68), .CopyingTeamBonuses)
            .CopyingTeamUsage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(69), .CopyingTeamUsage)
            .CopyingTeamCostModifier = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(70), .CopyingTeamCostModifier)
            .InventionTeamName = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(71), .InventionTeamName)
            .InventionTeamBonuses = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(72), .InventionTeamBonuses)
            .InventionTeamUsage = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(73), .InventionTeamUsage)
            .InventionTeamCostModifier = GetColumnNumber(chkLstBoxColumns.GetItemCheckState(74), .InventionTeamCostModifier)

            ' Now in case something was removed, we want to update the indicies
            With UserManufacturingTabColumnSettings
                For i = 0 To ColumnPositions.Count - 1
                    ColumnPositions(i) = ""
                Next

                With UserManufacturingTabColumnSettings
                    ColumnPositions(.ItemCategory) = ProgramSettings.ItemCategoryColumnName
                    ColumnPositions(.ItemGroup) = ProgramSettings.ItemGroupColumnName
                    ColumnPositions(.ItemName) = ProgramSettings.ItemNameColumnName
                    ColumnPositions(.Owned) = ProgramSettings.OwnedColumnName
                    ColumnPositions(.Tech) = ProgramSettings.TechColumnName
                    ColumnPositions(.BPME) = ProgramSettings.BPMEColumnName
                    ColumnPositions(.BPTE) = ProgramSettings.BPTEColumnName
                    ColumnPositions(.Inputs) = ProgramSettings.InputsColumnName
                    ColumnPositions(.Compared) = ProgramSettings.ComparedColumnName
                    ColumnPositions(.Runs) = ProgramSettings.RunsColumnName
                    ColumnPositions(.ProductionLines) = ProgramSettings.ProductionLinesColumnName
                    ColumnPositions(.LaboratoryLines) = ProgramSettings.LaboratoryLinesColumnName
                    ColumnPositions(.InventionCost) = ProgramSettings.InventionCostColumnName
                    ColumnPositions(.CopyCost) = ProgramSettings.CopyCostColumnName
                    ColumnPositions(.TotalManufacturingCost) = ProgramSettings.TotalManufacturingCostColumnName
                    ColumnPositions(.Taxes) = ProgramSettings.TaxesColumnName
                    ColumnPositions(.BrokerFees) = ProgramSettings.BrokerFeesColumnName
                    ColumnPositions(.BPProductionTime) = ProgramSettings.BPProductionTimeColumnName
                    ColumnPositions(.TotalProductionTime) = ProgramSettings.TotalProductionTimeColumnName
                    ColumnPositions(.ItemMarketPrice) = ProgramSettings.ItemMarketPriceColumnName
                    ColumnPositions(.Profit) = ProgramSettings.ProfitColumnName
                    ColumnPositions(.ProfitPercentage) = ProgramSettings.ProfitPercentageColumnName
                    ColumnPositions(.IskperHour) = ProgramSettings.IskperHourColumnName
                    ColumnPositions(.SVR) = ProgramSettings.SVRColumnName
                    ColumnPositions(.TotalCost) = ProgramSettings.TotalCostColumnName
                    ColumnPositions(.BaseJobCost) = ProgramSettings.BaseJobCostColumnName
                    ColumnPositions(.ManufacturingJobFee) = ProgramSettings.ManufacturingJobFeeColumnName
                    ColumnPositions(.ManufacturingFacilityName) = ProgramSettings.ManufacturingFacilityNameColumnName
                    ColumnPositions(.ManufacturingFacilitySystem) = ProgramSettings.ManufacturingFacilitySystemColumnName
                    ColumnPositions(.ManufacturingFacilitySystemIndex) = ProgramSettings.ManufacturingFacilitySystemIndexColumnName
                    ColumnPositions(.ManufacturingFacilityTax) = ProgramSettings.ManufacturingFacilityTaxColumnName
                    ColumnPositions(.ManufacturingFacilityRegion) = ProgramSettings.ManufacturingFacilityRegionColumnName
                    ColumnPositions(.ManufacturingFacilityMEBonus) = ProgramSettings.ManufacturingFacilityMEBonusColumnName
                    ColumnPositions(.ManufacturingFacilityTEBonus) = ProgramSettings.ManufacturingFacilityTEBonusColumnName
                    ColumnPositions(.ManufacturingFacilityUsage) = ProgramSettings.ManufacturingFacilityUsageColumnName
                    ColumnPositions(.ComponentFacilityName) = ProgramSettings.ComponentFacilityNameColumnName
                    ColumnPositions(.ComponentFacilitySystem) = ProgramSettings.ComponentFacilitySystemColumnName
                    ColumnPositions(.ComponentFacilitySystemIndex) = ProgramSettings.ComponentFacilitySystemIndexColumnName
                    ColumnPositions(.ComponentFacilityTax) = ProgramSettings.ComponentFacilityTaxColumnName
                    ColumnPositions(.ComponentFacilityRegion) = ProgramSettings.ComponentFacilityRegionColumnName
                    ColumnPositions(.ComponentFacilityMEBonus) = ProgramSettings.ComponentFacilityMEBonusColumnName
                    ColumnPositions(.ComponentFacilityTEBonus) = ProgramSettings.ComponentFacilityTEBonusColumnName
                    ColumnPositions(.ComponentFacilityUsage) = ProgramSettings.ComponentFacilityUsageColumnName
                    ColumnPositions(.CopyingFacilityName) = ProgramSettings.CopyingFacilityNameColumnName
                    ColumnPositions(.CopyingFacilitySystem) = ProgramSettings.CopyingFacilitySystemColumnName
                    ColumnPositions(.CopyingFacilitySystemIndex) = ProgramSettings.CopyingFacilitySystemIndexColumnName
                    ColumnPositions(.CopyingFacilityTax) = ProgramSettings.CopyingFacilityTaxColumnName
                    ColumnPositions(.CopyingFacilityRegion) = ProgramSettings.CopyingFacilityRegionColumnName
                    ColumnPositions(.CopyingFacilityMEBonus) = ProgramSettings.CopyingFacilityMEBonusColumnName
                    ColumnPositions(.CopyingFacilityTEBonus) = ProgramSettings.CopyingFacilityTEBonusColumnName
                    ColumnPositions(.CopyingFacilityUsage) = ProgramSettings.CopyingFacilityUsageColumnName
                    ColumnPositions(.InventionFacilityName) = ProgramSettings.InventionFacilityNameColumnName
                    ColumnPositions(.InventionFacilitySystem) = ProgramSettings.InventionFacilitySystemColumnName
                    ColumnPositions(.InventionFacilitySystemIndex) = ProgramSettings.InventionFacilitySystemIndexColumnName
                    ColumnPositions(.InventionFacilityTax) = ProgramSettings.InventionFacilityTaxColumnName
                    ColumnPositions(.InventionFacilityRegion) = ProgramSettings.InventionFacilityRegionColumnName
                    ColumnPositions(.InventionFacilityMEBonus) = ProgramSettings.InventionFacilityMEBonusColumnName
                    ColumnPositions(.InventionFacilityTEBonus) = ProgramSettings.InventionFacilityTEBonusColumnName
                    ColumnPositions(.InventionFacilityUsage) = ProgramSettings.InventionFacilityUsageColumnName
                    ColumnPositions(.ManufacturingTeamName) = ProgramSettings.ManufacturingTeamNameColumnName
                    ColumnPositions(.ManufacturingTeamBonuses) = ProgramSettings.ManufacturingTeamBonusesColumnName
                    ColumnPositions(.ManufacturingTeamUsage) = ProgramSettings.ManufacturingTeamUsageColumnName
                    ColumnPositions(.ManufacturingTeamCostModifier) = ProgramSettings.ManufacturingTeamCostModifierColumnName
                    ColumnPositions(.ComponentTeamName) = ProgramSettings.ComponentTeamNameColumnName
                    ColumnPositions(.ComponentTeamBonuses) = ProgramSettings.ComponentTeamBonusesColumnName
                    ColumnPositions(.ComponentTeamUsage) = ProgramSettings.ComponentTeamUsageColumnName
                    ColumnPositions(.ComponentTeamCostModifier) = ProgramSettings.ComponentTeamCostModifierColumnName
                    ColumnPositions(.CopyingTeamName) = ProgramSettings.CopyingTeamNameColumnName
                    ColumnPositions(.CopyingTeamBonuses) = ProgramSettings.CopyingTeamBonusesColumnName
                    ColumnPositions(.CopyingTeamUsage) = ProgramSettings.CopyingTeamUsageColumnName
                    ColumnPositions(.CopyingTeamCostModifier) = ProgramSettings.CopyingTeamCostModifierColumnName
                    ColumnPositions(.InventionTeamName) = ProgramSettings.InventionTeamNameColumnName
                    ColumnPositions(.InventionTeamBonuses) = ProgramSettings.InventionTeamBonusesColumnName
                    ColumnPositions(.InventionTeamUsage) = ProgramSettings.InventionTeamUsageColumnName
                    ColumnPositions(.InventionTeamCostModifier) = ProgramSettings.InventionTeamCostModifierColumnName
                End With

                ' Reset the first one with nothing since the first column is empty
                ColumnPositions(0) = "BPTypeID"

                ' Now get the total number of columns in the list we want to see
                For i = 1 To ColumnPositions.Count - 1
                    If ColumnPositions(i) <> "" Then
                        ColumnCount += 1
                    End If
                Next

                ' Init temp
                For i = 0 To TempColumns.Count - 1
                    TempColumns(i) = ""
                Next

                ' Now loop through the columns and update the positions
                For i = 1 To ColumnPositions.Count - 1
                    If ColumnPositions(i) <> "" Then
                        TempColumns(j) = ColumnPositions(i)
                        j += 1
                    Else
                        If i = UserManufacturingTabColumnSettings.OrderByColumn Then
                            ' They removed the column they sorted, so default to the first column since you must have 1
                            UserManufacturingTabColumnSettings.OrderByColumn = 1
                        End If
                    End If
                Next

                ColumnPositions = TempColumns

                ' Finally save the columns based on the current order
                With UserManufacturingTabColumnSettings
                    For i = 1 To ColumnPositions.Count - 1
                        Select Case ColumnPositions(i)
                            Case ProgramSettings.ItemCategoryColumnName
                                .ItemCategory = i
                            Case ProgramSettings.ItemGroupColumnName
                                .ItemGroup = i
                            Case ProgramSettings.ItemNameColumnName
                                .ItemName = i
                            Case ProgramSettings.OwnedColumnName
                                .Owned = i
                            Case ProgramSettings.TechColumnName
                                .Tech = i
                            Case ProgramSettings.BPMEColumnName
                                .BPME = i
                            Case ProgramSettings.BPTEColumnName
                                .BPTE = i
                            Case ProgramSettings.InputsColumnName
                                .Inputs = i
                            Case ProgramSettings.ComparedColumnName
                                .Compared = i
                            Case ProgramSettings.RunsColumnName
                                .Runs = i
                            Case ProgramSettings.ProductionLinesColumnName
                                .ProductionLines = i
                            Case ProgramSettings.LaboratoryLinesColumnName
                                .LaboratoryLines = i
                            Case ProgramSettings.InventionCostColumnName
                                .InventionCost = i
                            Case ProgramSettings.CopyCostColumnName
                                .CopyCost = i
                            Case ProgramSettings.TotalManufacturingCostColumnName
                                .TotalManufacturingCost = i
                            Case ProgramSettings.TaxesColumnName
                                .Taxes = i
                            Case ProgramSettings.BrokerFeesColumnName
                                .BrokerFees = i
                            Case ProgramSettings.BPProductionTimeColumnName
                                .BPProductionTime = i
                            Case ProgramSettings.TotalProductionTimeColumnName
                                .TotalProductionTime = i
                            Case ProgramSettings.ItemMarketPriceColumnName
                                .ItemMarketPrice = i
                            Case ProgramSettings.ProfitColumnName
                                .Profit = i
                            Case ProgramSettings.ProfitPercentageColumnName
                                .ProfitPercentage = i
                            Case ProgramSettings.IskperHourColumnName
                                .IskperHour = i
                            Case ProgramSettings.SVRColumnName
                                .SVR = i
                            Case ProgramSettings.TotalCostColumnName
                                .TotalCost = i
                            Case ProgramSettings.BaseJobCostColumnName
                                .BaseJobCost = i
                            Case ProgramSettings.ManufacturingJobFeeColumnName
                                .ManufacturingJobFee = i
                            Case ProgramSettings.ManufacturingFacilityNameColumnName
                                .ManufacturingFacilityName = i
                            Case ProgramSettings.ManufacturingFacilitySystemColumnName
                                .ManufacturingFacilitySystem = i
                            Case ProgramSettings.ManufacturingFacilitySystemIndexColumnName
                                .ManufacturingFacilitySystemIndex = i
                            Case ProgramSettings.ManufacturingFacilityTaxColumnName
                                .ManufacturingFacilityTax = i
                            Case ProgramSettings.ManufacturingFacilityRegionColumnName
                                .ManufacturingFacilityRegion = i
                            Case ProgramSettings.ManufacturingFacilityMEBonusColumnName
                                .ManufacturingFacilityMEBonus = i
                            Case ProgramSettings.ManufacturingFacilityTEBonusColumnName
                                .ManufacturingFacilityTEBonus = i
                            Case ProgramSettings.ManufacturingFacilityUsageColumnName
                                .ManufacturingFacilityUsage = i
                            Case ProgramSettings.ComponentFacilityNameColumnName
                                .ComponentFacilityName = i
                            Case ProgramSettings.ComponentFacilitySystemColumnName
                                .ComponentFacilitySystem = i
                            Case ProgramSettings.ComponentFacilitySystemIndexColumnName
                                .ComponentFacilitySystemIndex = i
                            Case ProgramSettings.ComponentFacilityTaxColumnName
                                .ComponentFacilityTax = i
                            Case ProgramSettings.ComponentFacilityRegionColumnName
                                .ComponentFacilityRegion = i
                            Case ProgramSettings.ComponentFacilityMEBonusColumnName
                                .ComponentFacilityMEBonus = i
                            Case ProgramSettings.ComponentFacilityTEBonusColumnName
                                .ComponentFacilityTEBonus = i
                            Case ProgramSettings.ComponentFacilityUsageColumnName
                                .ComponentFacilityUsage = i
                            Case ProgramSettings.CopyingFacilityNameColumnName
                                .CopyingFacilityName = i
                            Case ProgramSettings.CopyingFacilitySystemColumnName
                                .CopyingFacilitySystem = i
                            Case ProgramSettings.CopyingFacilitySystemIndexColumnName
                                .CopyingFacilitySystemIndex = i
                            Case ProgramSettings.CopyingFacilityTaxColumnName
                                .CopyingFacilityTax = i
                            Case ProgramSettings.CopyingFacilityRegionColumnName
                                .CopyingFacilityRegion = i
                            Case ProgramSettings.CopyingFacilityMEBonusColumnName
                                .CopyingFacilityMEBonus = i
                            Case ProgramSettings.CopyingFacilityTEBonusColumnName
                                .CopyingFacilityTEBonus = i
                            Case ProgramSettings.CopyingFacilityUsageColumnName
                                .CopyingFacilityUsage = i
                            Case ProgramSettings.InventionFacilityNameColumnName
                                .InventionFacilityName = i
                            Case ProgramSettings.InventionFacilitySystemColumnName
                                .InventionFacilitySystem = i
                            Case ProgramSettings.InventionFacilitySystemIndexColumnName
                                .InventionFacilitySystemIndex = i
                            Case ProgramSettings.InventionFacilityTaxColumnName
                                .InventionFacilityTax = i
                            Case ProgramSettings.InventionFacilityRegionColumnName
                                .InventionFacilityRegion = i
                            Case ProgramSettings.InventionFacilityMEBonusColumnName
                                .InventionFacilityMEBonus = i
                            Case ProgramSettings.InventionFacilityTEBonusColumnName
                                .InventionFacilityTEBonus = i
                            Case ProgramSettings.InventionFacilityUsageColumnName
                                .InventionFacilityUsage = i
                            Case ProgramSettings.ManufacturingTeamNameColumnName
                                .ManufacturingTeamName = i
                            Case ProgramSettings.ManufacturingTeamBonusesColumnName
                                .ManufacturingTeamBonuses = i
                            Case ProgramSettings.ManufacturingTeamUsageColumnName
                                .ManufacturingTeamUsage = i
                            Case ProgramSettings.ManufacturingTeamCostModifierColumnName
                                .ManufacturingTeamCostModifier = i
                            Case ProgramSettings.ComponentTeamNameColumnName
                                .ComponentTeamName = i
                            Case ProgramSettings.ComponentTeamBonusesColumnName
                                .ComponentTeamBonuses = i
                            Case ProgramSettings.ComponentTeamUsageColumnName
                                .ComponentTeamUsage = i
                            Case ProgramSettings.ComponentTeamCostModifierColumnName
                                .ComponentTeamCostModifier = i
                            Case ProgramSettings.CopyingTeamNameColumnName
                                .CopyingTeamName = i
                            Case ProgramSettings.CopyingTeamBonusesColumnName
                                .CopyingTeamBonuses = i
                            Case ProgramSettings.CopyingTeamUsageColumnName
                                .CopyingTeamUsage = i
                            Case ProgramSettings.CopyingTeamCostModifierColumnName
                                .CopyingTeamCostModifier = i
                            Case ProgramSettings.InventionTeamNameColumnName
                                .InventionTeamName = i
                            Case ProgramSettings.InventionTeamBonusesColumnName
                                .InventionTeamBonuses = i
                            Case ProgramSettings.InventionTeamUsageColumnName
                                .InventionTeamUsage = i
                            Case ProgramSettings.InventionTeamCostModifierColumnName
                                .InventionTeamCostModifier = i
                        End Select
                    Next
                End With
            End With
        End With

    End Sub

    Private Sub chkLstBoxColumns_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles chkLstBoxColumns.SelectedIndexChanged

        If SelectedIndex <> chkLstBoxColumns.SelectedIndex Then
            SelectedIndex = chkLstBoxColumns.SelectedIndex

            If chkLstBoxColumns.GetItemChecked(chkLstBoxColumns.SelectedIndex) Then
                ' Uncheckit
                chkLstBoxColumns.SetItemChecked(chkLstBoxColumns.SelectedIndex, False)
            Else
                ' Checkit
                chkLstBoxColumns.SetItemChecked(chkLstBoxColumns.SelectedIndex, True)
            End If

        End If

    End Sub

    Private Sub btnDefaults_Click(sender As System.Object, e As System.EventArgs) Handles btnDefaults.Click
        UserManufacturingTabColumnSettings = AllSettings.SetDefaultManufacturingTabColumnSettings()
        Call ShowList()
        chkLstBoxColumns.Update()
    End Sub

End Class