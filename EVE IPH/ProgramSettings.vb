
Imports System.Xml
Imports System.IO
Imports System.Xml.XPath
Imports System.Text.RegularExpressions
Imports System.Data.SQLite

Public Module SettingsVariables

    ' All settings
    Public AllSettings As New ProgramSettings
    ' User Settings
    Public UserApplicationSettings As ApplicationSettings
    ' Tower Cost settings
    Public SelectedTower As PlayerOwnedStationSettings
    ' BP Tab Settings
    Public UserBPTabSettings As BPTabSettings
    ' Manufacturing
    Public UserManufacturingTabSettings As ManufacturingTabSettings
    ' Datacores
    Public UserDCTabSettings As DataCoreTabSettings
    ' Reactions Tab
    Public UserReactionTabSettings As ReactionsTabSettings
    ' Update Prices Tab Settings
    Public UserUpdatePricesTabSettings As UpdatePriceTabSettings
    ' Mining Tab Settings
    Public UserMiningTabSettings As MiningTabSettings
    ' Industry Job Column Settings
    Public UserIndustryJobsColumnSettings As IndustryJobsColumnSettings
    ' Manufacturing Tab Column Settings
    Public UserManufacturingTabColumnSettings As ManufacturingTabColumnSettings
    ' Shopping List settings
    Public UserShoppingListSettings As ShoppingListSettings
    ' Industry Flip Belt Settings
    Public UserIndustryFlipBeltSettings As IndustryFlipBeltSettings
    ' and the five belts
    Public UserIndustryFlipBeltOreCheckSettings1 As IndustryBeltOreChecks
    Public UserIndustryFlipBeltOreCheckSettings2 As IndustryBeltOreChecks
    Public UserIndustryFlipBeltOreCheckSettings3 As IndustryBeltOreChecks
    Public UserIndustryFlipBeltOreCheckSettings4 As IndustryBeltOreChecks
    Public UserIndustryFlipBeltOreCheckSettings5 As IndustryBeltOreChecks
    ' Asset windows - multiple
    Public UserAssetWindowDefaultSettings As AssetWindowSettings
    Public UserAssetWindowShoppingListSettings As AssetWindowSettings

End Module

Public Class ProgramSettings

    ' Default Tower Settings
    Public Const DefaultTowerName As String = None
    Public Const DefaultTowerRaceID As Integer = 0
    Public Const DefaultCostperHour As Integer = 0
    Public Const DefaultMECostperSecond As Integer = 0
    Public Const DefaultTECostperSecond As Integer = 0
    Public Const DefaultInventionCostperSecond As Integer = 0
    Public Const DefaultCopyCostperSecond As Integer = 0
    Public Const DefaultTowerType As String = "Standard"
    Public Const DefaultTowerSize As String = "Large"
    Public Const DefaultFuelBlockBuild As Boolean = False
    Public Const DefaultNumAdvLabs As Integer = 0
    Public Const DefaultNumMobileLabs As Integer = 0
    Public Const DefaultNumHyasyodaLabs As Integer = 0
    Public Const DefaultCharterCost As Double = 2500.0

    ' Application Setting Defaults
    Public MBeanCounterName As String = "Zainou 'Beancounter' Industry BX-80" ' Manufacturing time
    Public RBeanCounterName As String = "Zainou 'Beancounter' Reprocessing RX-80" ' Refining waste
    Public CBeanCounterName As String = "Zainou 'Beancounter' Science SC-80" ' Copy time

    Public DefaultCheckUpdatesOnStart As Boolean = True
    Public DefaultAllowSkillOverride As Boolean = False
    Public DefaultDataExportFormat As String = "Default"
    Public DefaultShowToolTips As Boolean = True
    Public DefaultLoadAssetsonStartup As Boolean = True
    Public DefaultLoadBPsonStartup As Boolean = True
    Public DefaultRefreshTeamCRESTDataonStartup As Boolean = True
    Public DefaultRefreshMarketCRESTDataonStartup As Boolean = True
    Public DefaultRefreshFacilityCRESTDataonStartup As Boolean = True
    Public DefaultDisableSound As Boolean = False
    Public DefaultDNMarkInlineasOwned As Boolean = False

    Public DefaultBuildBaseInstall As Double = 1000
    Public DefaultBuildBaseHourly As Double = 333
    Public DefaultBuildStandingDiscount As Double = 0.015
    Public DefaultBuildStandingSurcharge As Double = 0.005

    Public DefaultInventBaseInstall As Double = 10000
    Public DefaultInventBaseHourly As Double = 416.67
    Public DefaultInventStandingDiscount As Double = 0.015
    Public DefaultInventStandingSurcharge As Double = 0.005

    Public DefaultBuildCorpStanding As Double = 5.0 ' Corp standing of where this blueprint will be made
    Public DefaultInventCorpStanding As Double = 5.0 ' Corp standing of where this blueprint will be invented
    Public DefaultBrokerCorpStanding As Double = 5.0 ' Corp standing of where this blueprint will be sold
    Public DefaultBrokerFactionStanding As Double = 5.0 ' Faction standing of where this blueprint will be sold (for Broker calc)
    Public DefaultRefineCorpStanding As Double = 6.67 ' Corp standing for use of refining

    Public DefaultIncludeCopyTimes As Boolean = False ' If we include copy times in IPH calcs for invention
    Public DefaultIncludeInventionTimes As Boolean = False ' If we include invention times in IPH calcs for invention
    Public DefaultIncludeRETimes As Boolean = False ' If we include RE times in IPH calcs for RE

    Public DefaultEstimateCopyCost As Boolean = False ' Estimate copy costs for invention BPC's
    Public DefaultCopySlotModifier As String = "1.0" ' The default copy slot modifier for T1 BPC copies to use in invention
    Public DefaultInventionSlotModifier As String = "1.0" ' Default invention time
    Public DefaultBuildSlotModifier As String = "1.0" ' Default build time for production
    Public DefaultRefiningEfficency As Double = 0.5 ' Default refining equipment

    Public DefaultRefineTax As Double = 0.05 ' Default tax rate

    Public DefaultCheckBuildBuy As Boolean = False
    Public DefaultLinkTeamstoFacilitySystems As Boolean = False

    Public DefaultSettingME As Integer = 0
    Public DefaultSettingTE As Integer = 0

    Public DefaultDisableSVR As Boolean = False
    Public DefaultSuggestBuildBPNotOwned As Boolean = True ' If the bp is not owned, default to suggesting they build the item anyway

    ' For shopping list
    Public DefaultShopListIncludeInventMats As Boolean = True
    Public DefaultShopListIncludeREMats As Boolean = True

    ' If the user has no implants
    Public DefaultImplantValues As Double = 0

    ' No team
    Public DefaultTeamID As Long = 0

    ' Default Facilities - all Jita 4-4 except for RE, which will be a POS in Jita 
    Public DefaultManufacturingFacilityID As Long = 60003760
    Public DefaultManufacturingFacility As String = "Jita IV - Moon 4 - Caldari Navy Assembly Plant"
    Public DefaultManufacturingFacilityType As String = StationFacility
    Public DefaultComponentManufacturingFacilityID As Long = 60003760
    Public DefaultComponentManufacturingFacility As String = "Jita IV - Moon 4 - Caldari Navy Assembly Plant"
    Public DefaultComponentManufacturingFacilityType As String = StationFacility
    Public DefaultCapitalComponentManufacturingFacilityID As Long = 60003760
    Public DefaultCapitalComponentManufacturingFacility As String = "Jita IV - Moon 4 - Caldari Navy Assembly Plant"
    Public DefaultCapitalComponentManufacturingFacilityType As String = StationFacility
    Public DefaultCapitalManufacturingFacilityID As Long = 60003043
    Public DefaultCapitalManufacturingFacility As String = "Akora VI - Moon 7 - Expert Housing Production Plant"
    Public DefaultCapitalManufacturingFacilityType As String = StationFacility
    Public DefaultCopyFacilityID As Long = 60001786 ' Wous
    Public DefaultCopyFacility As String = "Wuos VI - Zainou Biotech Research Center"
    Public DefaultCopyFacilityType As String = StationFacility
    Public DefaultInventionFacilityID As Long = 60001786 ' Wous
    Public DefaultInventionFacility As String = "Wuos VI - Zainou Biotech Research Center"
    Public DefaultInventionFacilityType As String = StationFacility
    Public DefaultT3InventionFacilityID As Long = 24567
    Public DefaultT3InventionFacility As String = "Experimental Laboratory"
    Public DefaultT3InventionFacilityType As String = POSFacility
    Public DefaultT3CruiserManufacturingFacilityID As Long = 30389 ' If we are manufacturing a T3 item, then default to the subsystem array in a pos
    Public DefaultT3CruiserManufacturingFacility As String = "Subsystem Assembly Array"
    Public DefaultT3CruiserManufacturingFacilityType As String = POSFacility
    Public DefaultT3DestroyerManufacturingFacilityID As Long = 24653 ' If we are manufacturing a T3 item, then default to the subsystem array in a pos
    Public DefaultT3DestroyerManufacturingFacility As String = "Advanced Small Ship Assembly Array"
    Public DefaultT3DestroyerManufacturingFacilityType As String = POSFacility
    Public DefaultSubsystemManufacturingFacilityID As Long = 30389 ' If we are manufacturing a T3 item, then default to the subsystem array in a pos
    Public DefaultSubsystemManufacturingFacility As String = "Subsystem Assembly Array"
    Public DefaultSubsystemManufacturingFacilityType As String = POSFacility
    Public DefaultSuperManufacturingFacilityID As Long = 24575 ' If we are manufacturing a super, then default to the supercaptial array in a pos
    Public DefaultSuperManufacturingFacility As String = "Supercapital Ship Assembly Array"
    Public DefaultSuperManufacturingFacilityType As String = POSFacility
    Public DefaultBoosterManufacturingFacilityID As Long = 25305 ' Drug lab in a pos
    Public DefaultBoosterManufacturingFacility As String = "Drug Lab"
    Public DefaultBoosterManufacturingFacilityType As String = POSFacility

    Public DefaultPOSFuelBlockManufacturingFacilityID As Long = 24660
    Public DefaultPOSFuelBlockManufacturingFacility As String = "Component Assembly Array" ' Component array in a pos for fuel blocks
    Public DefaultPOSFuelBlockManufacturingFacilityType As String = POSFacility

    Public DefaultPOSLargeShipManufacturingFacilityID As Long = 29613
    Public DefaultPOSLargeShipManufacturingFacility As String = "Large Ship Assembly Array" ' Large Ship assembly array in a pos for Large Ships
    Public DefaultPOSLargeShipManufacturingFacilityType As String = POSFacility

    Public DefaultPOSModuleManufacturingFacilityID As Long = 13780
    Public DefaultPOSModuleManufacturingFacility As String = "Equipment Assembly Array" ' Equipment assembly array in a pos for all modules
    Public DefaultPOSModuleManufacturingFacilityType As String = POSFacility

    Public FacilityDefaultMM As Double = 1
    Public FacilityDefaultTM As Double = 1
    Public DefalutFacilityCM As Double = 1
    Public FacilityDefaultTax As Double = 0.1

    ' For POS data (T3 and general pos)
    Public FacilityDefaultSolarSystemID As Long = 30000142
    Public FacilityDefaultSolarSystem As String = "Jita"
    Public FacilityDefaultRegionID As Long = 10000002
    Public FacilityDefaultRegion As String = "The Forge"

    ' For Booster and super pos production
    Public DefaultNullFacilitySolarSystemID As Long = 30003713
    Public DefaultNullFacilitySolarSystem As String = "G7AQ-7"
    Public DefaultNullFacilityRegionID As Long = 10000047
    Public DefaultNullFacilityRegion As String = "Providence"

    Public FacilityDefaultActivityCostperSecond As Double = 0
    Public FacilityDefaultIncludeUsage As Boolean = True
    Public FacilityDefaultIncludeCost As Boolean = False ' Only for Invention, Copy, and RE so let this get set 
    Public FacilityDefaultIncludeTime As Boolean = False ' Only for Invention, Copy, and RE so let this get set 

    ' Set here, but use in Update Prices - 6 hours to refresh prices
    Public DefaultEVECentralRefreshInterval As Integer = 6

    ' BP Tab Default settings
    Public DefaultBPTechChecks As Boolean = True
    Public DefaultSizeChecks As Boolean = False
    Public DefaultBPSelectionType As String = "All"
    Public DefaultBPIncludeFees As Boolean = True
    Public DefaultBPIncludeTaxes As Boolean = True
    Public DefaultBPIncludeUsage As Boolean = True
    Public DefaultBPIgnoreChecks As Boolean = False
    Public DefaultBPPricePerUnit As Boolean = False
    Public DefaultBPIncludeInventionTime As Boolean = False
    Public DefaultBPIncludeInventionCost As Boolean = False
    Public DefaultBPIncludeCopyTime As Boolean = False
    Public DefaultBPIncludecopyCost As Boolean = False
    Public DefaultBPIncludeT3Cost As Boolean = False
    Public DefaultBPIncludeT3Time As Boolean = False
    Public DefaultBPProductionLines As Integer = 1
    Public DefaultBPLaboratoryLines As Integer = 1
    Public DefaultBPRELines As Integer = 1

    ' Update Prices Default Settings
    Public DefaultPriceChecks As Boolean = False
    Public DefaultPriceImportPriceType As String = "Minimum Sell"
    Public DefaultPriceSystem As String = "Jita"
    Public DefaultPriceRegion As String = ""
    Public DefaultPriceRawMatsCombo As String = "Max Buy"
    Public DefaultPriceItemsCombo As String = "Min Sell"

    ' Default Manufacturing Tab
    Public DefaultBlueprintType As String = "All Blueprints"
    Public DefaultCheckTech1 As Boolean = True
    Public DefaultCheckTech2 As Boolean = True
    Public DefaultCheckTech3 As Boolean = True
    Public DefaultCheckTechStoryline As Boolean = True
    Public DefaultCheckTechNavy As Boolean = True
    Public DefaultCheckTechPirate As Boolean = True
    Public DefaultItemTypeFilter As String = "All Types"
    Public DefaultTextItemFilter As String = ""
    Public DefaultCheckBPTypeShips As Boolean = True
    Public DefaultCheckBPTypeDrones As Boolean = True
    Public DefaultCheckBPTypeComponents As Boolean = True
    Public DefaultCheckBPTypeStructures As Boolean = True
    Public DefaultCheckBPTypeTools As Boolean = True
    Public DefaultCheckBPTypeModules As Boolean = True
    Public DefaultCheckBPTypeAmmoCharges As Boolean = True
    Public DefaultCheckBPTypeRigs As Boolean = True
    Public DefaultCheckBPTypeSubsystems As Boolean = True
    Public DefaultCheckBPTypeBoosters As Boolean = True
    Public DefaultCheckBPTypeDeployables As Boolean = True
    Public DefaultCheckBPTypeCelestials As Boolean = True
    Public DefaultCheckBPTypeStationParts As Boolean = True
    Public DefaultAveragePriceDuration As String = "7"
    Public DefaultCheckDecryptorNone As Boolean = True
    Public DefaultCheckDecryptor06 As Boolean = False
    Public DefaultCheckDecryptor09 As Boolean = False
    Public DefaultCheckDecryptor10 As Boolean = False
    Public DefaultCheckDecryptor11 As Boolean = False
    Public DefaultCheckDecryptor12 As Boolean = False
    Public DefaultCheckDecryptor15 As Boolean = False
    Public DefaultCheckDecryptor18 As Boolean = False
    Public DefaultCheckDecryptor19 As Boolean = False
    Public DefaultCheckDecryptorUseforT2 As Boolean = True
    Public defaultCheckDecryptorUseforT3 As Boolean = True
    Public DefaultCheckIgnoreInventionRE As Boolean = False
    Public DefaultCheckRelicWrecked As Boolean = True
    Public DefaultCheckRelicIntact As Boolean = False
    Public DefaultCheckRelicMalfunction As Boolean = False
    Public DefaultCheckOnlyBuild As Boolean = False
    Public DefaultCheckOnlyInvent As Boolean = False
    Public DefaultCheckOnlyRE As Boolean = False
    Public DefaultCheckIncludeTaxes As Boolean = True
    Public DefaultIncludeBrokersFees As Boolean = True
    Public DefaultCheckIncludeUsage As Boolean = True
    Public DefaultCheckRaceAmarr As Boolean = True
    Public DefaultCheckRaceCaldari As Boolean = True
    Public DefaultCheckRaceGallente As Boolean = True
    Public DefaultCheckRaceMinmatar As Boolean = True
    Public DefaultCheckRacePirate As Boolean = True
    Public DefaultCheckRaceOther As Boolean = True
    Public DefaultSortBy As String = "IPH"
    Public DefaultPriceCompare As String = "Compare All"
    Public DefaultCheckIncludeT2Owned As Boolean = True
    Public DefaultCheckIncludeT3Owned As Boolean = True
    Public DefaultIgnoreSVRThreshold As Double = 0.0
    Public DefaultCheckSVRIncludeNull As Boolean = True
    Public DefaultSVRRegion As String = "The Forge"
    Public DefaultCalcProductionLines As Integer = 1
    Public DefaultCalcLaboratoryLines As Integer = 1
    Public DefaultCalcRuns As Integer = 1
    Public DefaultCalcSizeChecks As Boolean = False
    Public DefaultCheckT3Destroyers As Boolean = False
    Public DefaultCheckCapComponents As Boolean = False

    ' Datacore Default Settings
    Public DefaultDCPricesFrom As String = "Updated Prices"
    Public DefaultDCCheckHighSec As Boolean = True
    Public DefaultDCCheckLowNullSec As Boolean = False
    Public DefaultDCIncludeAgentsCantUse As Boolean = False
    Public DefaultDCAgentsInRegion As String = "All Regions"
    Public DefaultDCSovCheck As Boolean = True

    ' Datacores For these, use the users settings
    Public DefaultConnections As Integer = -1
    Public DefaultNegotiation As Integer = -1
    Public DefaultResearchProjMgt As Integer = -1
    Public DefaultCorpStanding As Integer = -1
    Public DefaultCorpStandingChecked As Integer = -1
    Public DefaultSkillLevel As Integer = -1
    Public DefaultSkillLevelChecked As Integer = -1

    ' Datacore setting array sizes
    Public NumberofDCSettingsSkillRecords As Integer = 16
    Public NumberofDCSettingsCorpRecords As Integer = 12

    ' Reactions Default Settings
    Public DefaultReactPOSFuelCost As Double = 500000.0
    Public DefaultReactCheckTaxes As Boolean = True
    Public DefaultReactCheckFees As Boolean = True
    Public DefaultReactItemChecks As Boolean = False
    Public DefaultReactNumPOS As Integer = 1

    ' Mining Default Settings
    Public DefaultMiningOreType As String = "Ore"
    Public DefaultMiningCheckHighYieldOres As Boolean = False
    Public DefaultMiningCheckHighSecOres As Boolean = True
    Public DefaultMiningCheckLowSecOres As Boolean = False
    Public DefaultMiningCheckNullSecOres As Boolean = False
    Public DefaultMiningCheckSovAmarr As Boolean = True
    Public DefaultMiningCheckSovCaldari As Boolean = True
    Public DefaultMiningCheckSovGallente As Boolean = True
    Public DefaultMiningCheckSovMinmatar As Boolean = True
    Public DefaultMiningCheckSovWormhole As Boolean = True
    Public DefaultMiningCheckSovC1 As Boolean = True
    Public DefaultMiningCheckSovC2 As Boolean = True
    Public DefaultMiningCheckSovC3 As Boolean = True
    Public DefaultMiningCheckSovC4 As Boolean = True
    Public DefaultMiningCheckSovC5 As Boolean = True
    Public DefaultMiningCheckSovC6 As Boolean = True
    Public DefaultMiningCheckIncludeFees As Boolean = True
    Public DefaultMiningCheckIncludeTaxes As Boolean = True
    Public DefaultMiningCheckIncludeJumpFuelCosts As Boolean = False
    Public DefaultMiningTotalJumpFuelCost As Integer = 0
    Public DefaultMiningTotalJumpFuelM3 As Integer = 1
    Public DefaultMiningJumpCompressedOre As Boolean = True
    Public DefaultMiningJumpMinerals As Boolean = False
    Public DefaultMiningMiningShip As String = "" ' Keep this blank so that it will default to a ship for them, if they have the skills
    Public DefaultMiningIceMiningShip As String = "" ' Keep this blank so that it will default to a ship for them, if they have the skills
    Public DefaultMiningGasMiningShip As String = ""
    Public DefaultMiningOreStrip As String = "" ' Keep blank to set max possible strip/miner they can use
    Public DefaultMiningIceStrip As String = "" ' Keep blank so they can set the max possible ice strip
    Public DefaultMiningGasHarvester As String = ""
    Public DefaultMiningNumOreMiners As Integer = 0
    Public DefaultMiningNumIceMiners As Integer = 0
    Public DefaultMiningNumGasHarvesters As Integer = 0
    Public DefaultMiningOreUpgrade As String = None
    Public DefaultMiningIceUpgrade As String = None
    Public DefaultMiningGasUpgrade As String = None
    Public DefaultMiningNumOreUpgrades As Integer = 0
    Public DefaultMiningNumIceUpgrades As Integer = 0
    Public DefaultMiningNumGasUpgrades As Integer = 0
    Public DefaultMiningMichiiImplant As Boolean = False
    Public DefaultMiningT2Crystals As Boolean = False
    Public DefaultMiningOreImplant As String = None
    Public DefaultMiningIceImplant As String = None
    Public DefaultMiningGasImplant As String = None
    Public DefaultMiningCheckUseHauler As Boolean = True
    Public DefaultMiningRoundTripMin As Integer = 1
    Public DefaultMiningRoundTripSec As Integer = 0
    Public DefaultMiningHaulerm3 As Integer = 0
    Public DefaultMiningCheckUseFleetBooster As Boolean = False
    Public DefaultMiningBoosterShip As String = "Other"
    Public DefaultMiningBoosterShipSkill As Integer = 0
    Public DefaultMiningMiningFormanSkill As Integer = 0
    Public DefaultMiningMiningDirectorSkill As Integer = 0
    Public DefaultMiningWarfareLinkSpecSkill As Integer = 0
    Public DefaultMiningCheckMineForemanLaserOpBoost As Integer = 0
    Public DefaultMiningCheckMiningForemanMindLink As Boolean = False
    Public DefaultMiningRefineCorpTax As Double = 0.05
    Public DefaultMiningRorqDeployed As Boolean = True
    Public DefaultMiningDroneM3perHour As Double = 0.0
    Public DefaultMiningRefineOre As Boolean = True
    Public DefaultIndustrialReconfig As Integer = 0
    Public DefaultMiningRig As Boolean = False

    ' Industry Jobs column settings
    Public DefaultJobState As Integer = 1
    Public DefaultTimeToComplete As Integer = 3
    Public DefaultActivity As Integer = 2
    Public DefaultStatus As Integer = 0
    Public DefaultStartTime As Integer = 0
    Public DefaultEndTime As Integer = 0
    Public DefaultCompletionTime As Integer = 0
    Public DefaultBlueprint As Integer = 4
    Public DefaultOutputItem As Integer = 5
    Public DefaultOutputItemType As Integer = 0
    Public DefaultInstallSolarSystem As Integer = 6
    Public DefaultInstallRegion As Integer = 7
    Public DefaultLicensedRuns As Integer = 0
    Public DefaultRuns As Integer = 0
    Public DefaultSuccessfulRuns As Integer = 0
    Public DefaultBlueprintLocation As Integer = 8
    Public DefaultOutputLocation As Integer = 9
    Public DefaultOrderType As String = "Ascending"
    Public DefaultViewJobType As String = "Personal"
    Public DefaultJobTimes As String = "Current Jobs"
    Public DefaultIndustryColumnWidth As Integer = 100
    Public DefaultOrderByColumn As Integer = 3

    ' Column Names for industry jobs viewer
    Public Const JobStateColumn As String = "Job State"
    Public Const TimetoCompleteColumn As String = "Time to Complete"
    Public Const ActivityColumn As String = "Activity"
    Public Const StatusColumn As String = "Status"
    Public Const StartTimeColumn As String = "Start Time"
    Public Const EndTimeColumn As String = "End Time"
    Public Const CompletionTimeColumn As String = "Completed Time"
    Public Const BlueprintColumn As String = "Blueprint"
    Public Const OutputItemColumn As String = "Output Item"
    Public Const OutputItemTypeColumn As String = "Output Item Type"
    Public Const InstallSolarSystemColumn As String = "Install System"
    Public Const InstallRegionColumn As String = "Install Region"
    Public Const LicensedRunsColumn As String = "Licensed Runs"
    Public Const RunsColumn As String = "Runs"
    Public Const SuccessfulRunsColumn As String = "Successful Runs"
    Public Const BlueprintLocationColumn As String = "Blueprint Location"
    Public Const OutputLocationColumn As String = "Output Location"

    ' Manufacturing Tab column settings
    Public DefaultMTItemCategory As Integer = 1
    Public DefaultMTItemGroup As Integer = 2
    Public DefaultMTItemName As Integer = 3
    Public DefaultMTOwned As Integer = 4
    Public DefaultMTTech As Integer = 5
    Public DefaultMTBPME As Integer = 6
    Public DefaultMTBPTE As Integer = 7
    Public DefaultMTInputs As Integer = 8
    Public DefaultMTCompared As Integer = 9
    Public DefaultMTRuns As Integer = 10
    Public DefaultMTProductionLines As Integer = 11
    Public DefaultMTLaboratoryLines As Integer = 12
    Public DefaultMTTotalInventionRECost As Integer = 13
    Public DefaultMTTotalCopyCost As Integer = 14
    Public DefaultMTTotalManufacturingCost As Integer = 15
    Public DefaultMTTaxes As Integer = 16
    Public DefaultMTBrokerFees As Integer = 17
    Public DefaultMTBPProductionTime As Integer = 18
    Public DefaultMTTotalProductionTime As Integer = 19
    Public DefaultMTItemMarketPrice As Integer = 20
    Public DefaultMTProfit As Integer = 21
    Public DefaultMTProfitPercentage As Integer = 22
    Public DefaultMTIskperHour As Integer = 23
    Public DefaultMTSVR As Integer = 24
    Public DefaultMTTotalCost As Integer = 25
    Public DefaultMTBaseJobCost As Integer = 26
    Public DefaultMTManufacturingJobFee As Integer = 27
    Public DefaultMTManufacturingFacilityName As Integer = 28
    Public DefaultMTManufacturingFacilitySystem As Integer = 29
    Public DefaultMTManufacturingFacilitySystemIndex As Integer = 30
    Public DefaultMTManufacturingFacilityTax As Integer = 31
    Public DefaultMTManufacturingFacilityRegion As Integer = 32
    Public DefaultMTManufacturingFacilityMEBonus As Integer = 33
    Public DefaultMTManufacturingFacilityTEBonus As Integer = 34
    Public DefaultMTManufacturingFacilityUsage As Integer = 35
    Public DefaultMTComponentFacilityName As Integer = 36
    Public DefaultMTComponentFacilitySystem As Integer = 37
    Public DefaultMTComponentFacilitySystemIndex As Integer = 38
    Public DefaultMTComponentFacilityTax As Integer = 39
    Public DefaultMTComponentFacilityRegion As Integer = 40
    Public DefaultMTComponentFacilityMEBonus As Integer = 41
    Public DefaultMTComponentFacilityTEBonus As Integer = 42
    Public DefaultMTComponentFacilityUsage As Integer = 43
    Public DefaultMTCopyingFacilityName As Integer = 44
    Public DefaultMTCopyingFacilitySystem As Integer = 45
    Public DefaultMTCopyingFacilitySystemIndex As Integer = 46
    Public DefaultMTCopyingFacilityTax As Integer = 47
    Public DefaultMTCopyingFacilityRegion As Integer = 48
    Public DefaultMTCopyingFacilityMEBonus As Integer = 49
    Public DefaultMTCopyingFacilityTEBonus As Integer = 50
    Public DefaultMTCopyingFacilityUsage As Integer = 51
    Public DefaultMTInventionREFacilityName As Integer = 52
    Public DefaultMTInventionREFacilitySystem As Integer = 53
    Public DefaultMTInventionREFacilitySystemIndex As Integer = 54
    Public DefaultMTInventionREFacilityTax As Integer = 55
    Public DefaultMTInventionREFacilityRegion As Integer = 56
    Public DefaultMTInventionREFacilityMEBonus As Integer = 57
    Public DefaultMTInventionREFacilityTEBonus As Integer = 58
    Public DefaultMTInventionREFacilityUsage As Integer = 59
    Public DefaultMTManufacturingTeamName As Integer = 60
    Public DefaultMTManufacturingTeamBonuses As Integer = 61
    Public DefaultMTManufacturingTeamUsage As Integer = 62
    Public DefaultMTManufacturingTeamCostModifier As Integer = 63
    Public DefaultMTComponentTeamName As Integer = 64
    Public DefaultMTComponentTeamBonuses As Integer = 65
    Public DefaultMTComponentTeamUsage As Integer = 66
    Public DefaultMTComponentTeamCostModifier As Integer = 67
    Public DefaultMTCopyingTeamName As Integer = 68
    Public DefaultMTCopyingTeamBonuses As Integer = 69
    Public DefaultMTCopyingTeamUsage As Integer = 70
    Public DefaultMTCopyingTeamCostModifier As Integer = 71
    Public DefaultMTInventionRETeamName As Integer = 72
    Public DefaultMTInventionRETeamBonuses As Integer = 73
    Public DefaultMTInventionRETeamUsage As Integer = 74
    Public DefaultMTInventionRETeamCostModifier As Integer = 75

    Public DefaultMTOrderType As String = "Ascending"
    Public DefaultMTColumnWidth As Integer = 100
    Public DefaultMTOrderByColumn As Integer = 3

    ' Column Names for manufacturing tab
    Public Const ItemCategoryColumnName As String = "Item Category"
    Public Const ItemGroupColumnName As String = "Item Group"
    Public Const ItemNameColumnName As String = "Item Name"
    Public Const OwnedColumnName As String = "Owned"
    Public Const TechColumnName As String = "Tech"
    Public Const BPMEColumnName As String = "ME"
    Public Const BPTEColumnName As String = "TE"
    Public Const InputsColumnName As String = "Inputs"
    Public Const ComparedColumnName As String = "Compared"
    Public Const RunsColumnName As String = "Runs"
    Public Const ProductionLinesColumnName As String = "Production Lines"
    Public Const LaboratoryLinesColumnName As String = "Laboratory Lines"
    Public Const TotalInventionRECostColumnName As String = "Total Invention / RE Cost"
    Public Const TotalCopyCostColumnName As String = "Total Copy Cost"
    Public Const TotalManufacturingCostColumnName As String = "Total Manufacturing Cost"
    Public Const TaxesColumnName As String = "Taxes"
    Public Const BrokerFeesColumnName As String = "BrokerFees"
    Public Const BPProductionTimeColumnName As String = "BP Production Time"
    Public Const TotalProductionTimeColumnName As String = "Total Production Time"
    Public Const ItemMarketPriceColumnName As String = "Item Market Price"
    Public Const ProfitColumnName As String = "Profit"
    Public Const ProfitPercentageColumnName As String = "Profit Percentage"
    Public Const IskperHourColumnName As String = "Isk per Hour"
    Public Const SVRColumnName As String = "SVR"
    Public Const TotalCostColumnName As String = "Total Cost"
    Public Const BaseJobCostColumnName As String = "Base Job Cost"
    Public Const ManufacturingJobFeeColumnName As String = "Manufacturing Job Fee"
    Public Const ManufacturingFacilityNameColumnName As String = "Manufacturing Facility Name"
    Public Const ManufacturingFacilitySystemColumnName As String = "Manufacturing Facility System"
    Public Const ManufacturingFacilitySystemIndexColumnName As String = "Manufacturing Facility System Index"
    Public Const ManufacturingFacilityTaxColumnName As String = "Manufacturing Facility Tax"
    Public Const ManufacturingFacilityRegionColumnName As String = "Manufacturing Facility Region"
    Public Const ManufacturingFacilityMEBonusColumnName As String = "Manufacturing Facility ME Bonus"
    Public Const ManufacturingFacilityTEBonusColumnName As String = "Manufacturing Facility TE Bonus"
    Public Const ManufacturingFacilityUsageColumnName As String = "Manufacturing Facility Usage"
    Public Const ComponentFacilityNameColumnName As String = "Component Facility Name"
    Public Const ComponentFacilitySystemColumnName As String = "Component Facility System"
    Public Const ComponentFacilitySystemIndexColumnName As String = "Component Facility System Index"
    Public Const ComponentFacilityTaxColumnName As String = "Component Facility Tax"
    Public Const ComponentFacilityRegionColumnName As String = "Component Facility Region"
    Public Const ComponentFacilityMEBonusColumnName As String = "Component Facility ME Bonus"
    Public Const ComponentFacilityTEBonusColumnName As String = "Component Facility TE Bonus"
    Public Const ComponentFacilityUsageColumnName As String = "Component Facility Usage"
    Public Const CopyingFacilityNameColumnName As String = "Copying Facility Name"
    Public Const CopyingFacilitySystemColumnName As String = "Copying Facility System"
    Public Const CopyingFacilitySystemIndexColumnName As String = "Copying Facility System Index"
    Public Const CopyingFacilityTaxColumnName As String = "Copying Facility Tax"
    Public Const CopyingFacilityRegionColumnName As String = "Copying Facility Region"
    Public Const CopyingFacilityMEBonusColumnName As String = "Copying Facility ME Bonus"
    Public Const CopyingFacilityTEBonusColumnName As String = "Copying Facility TE Bonus"
    Public Const CopyingFacilityUsageColumnName As String = "Copying Facility Usage"
    Public Const InventionREFacilityNameColumnName As String = "Invention / RE Facility Name"
    Public Const InventionREFacilitySystemColumnName As String = "Invention / RE Facility System"
    Public Const InventionREFacilitySystemIndexColumnName As String = "Invention / RE Facility System Index"
    Public Const InventionREFacilityTaxColumnName As String = "Invention / RE Facility Tax"
    Public Const InventionREFacilityRegionColumnName As String = "Invention / RE  Facility Region"
    Public Const InventionREFacilityMEBonusColumnName As String = "Invention / RE Facility ME Bonus"
    Public Const InventionREFacilityTEBonusColumnName As String = "Invention / RE Facility TE Bonus"
    Public Const InventionREFacilityUsageColumnName As String = "Invention / RE Facility Usage"
    Public Const ManufacturingTeamNameColumnName As String = "Manufacturing Team Name"
    Public Const ManufacturingTeamBonusesColumnName As String = "Manufacturing Team Bonuses"
    Public Const ManufacturingTeamUsageColumnName As String = "Manufacturing Team Usage"
    Public Const ManufacturingTeamCostModifierColumnName As String = "Manufacturing Team Cost Modifier"
    Public Const ComponentTeamNameColumnName As String = "Component Team Name"
    Public Const ComponentTeamBonusesColumnName As String = "Component Team Bonuses"
    Public Const ComponentTeamUsageColumnName As String = "Component Team Usage"
    Public Const ComponentTeamCostModifierColumnName As String = "Component Team Cost Modifier"
    Public Const CopyingTeamNameColumnName As String = "Copying Team Name"
    Public Const CopyingTeamBonusesColumnName As String = "Copying Team Bonuses"
    Public Const CopyingTeamUsageColumnName As String = "Copying Team Usage"
    Public Const CopyingTeamCostModifierColumnName As String = "Copying Team Cost Modifier"
    Public Const InventionRETeamNameColumnName As String = "Invention / RE Team Name"
    Public Const InventionRETeamBonusesColumnName As String = "Invention / RE Team Bonuses"
    Public Const InventionRETeamUsageColumnName As String = "Invention / RE Team Usage"
    Public Const InventionRETeamCostModifierColumnName As String = "Invention / RE Team Cost Modifier"

    ' Industry Flip Belt settings
    Private DefaultCycleTime As Double = 180
    Private Defaultm3perCycle As Double = 3000
    Private DefaultNumMiners As Integer = 1
    Private DefaultCompressOre As Boolean = False
    Private DefaultIPHperMiner As Boolean = False
    Private DefaultIncludeBrokerFees As Boolean = True
    Private DefaultIncludeTaxes As Boolean = True
    Private DefaultTruesec As String = ""

    ' Industry flip belt defaults
    Private DefaultPlagioclase As Boolean = True
    Private DefaultSpodumain As Boolean = True
    Private DefaultKernite As Boolean = True
    Private DefaultHedbergite As Boolean = True
    Private DefaultArkonor As Boolean = True
    Private DefaultBistot As Boolean = True
    Private DefaultPyroxeres As Boolean = True
    Private DefaultCrokite As Boolean = True
    Private DefaultJaspet As Boolean = True
    Private DefaultOmber As Boolean = True
    Private DefaultScordite As Boolean = True
    Private DefaultGneiss As Boolean = True
    Private DefaultVeldspar As Boolean = True
    Private DefaultHemorphite As Boolean = True
    Private DefaultDarkOchre As Boolean = True
    Private DefaultMercoxit As Boolean = True
    Private DefaultCrimsonArkonor As Boolean = True
    Private DefaultPrimeArkonor As Boolean = True
    Private DefaultTriclinicBistot As Boolean = True
    Private DefaultMonoclinicBistot As Boolean = True
    Private DefaultSharpCrokite As Boolean = True
    Private DefaultCrystallineCrokite As Boolean = True
    Private DefaultOnyxOchre As Boolean = True
    Private DefaultObsidianOchre As Boolean = True
    Private DefaultVitricHedbergite As Boolean = True
    Private DefaultGlazedHedbergite As Boolean = True
    Private DefaultVividHemorphite As Boolean = True
    Private DefaultRadiantHemorphite As Boolean = True
    Private DefaultPureJaspet As Boolean = True
    Private DefaultPristineJaspet As Boolean = True
    Private DefaultLuminousKernite As Boolean = True
    Private DefaultFieryKernite As Boolean = True
    Private DefaultAzurePlagioclase As Boolean = True
    Private DefaultRichPlagioclase As Boolean = True
    Private DefaultSolidPyroxeres As Boolean = True
    Private DefaultViscousPyroxeres As Boolean = True
    Private DefaultCondensedScordite As Boolean = True
    Private DefaultMassiveScordite As Boolean = True
    Private DefaultBrightSpodumain As Boolean = True
    Private DefaultGleamingSpodumain As Boolean = True
    Private DefaultConcentratedVeldspar As Boolean = True
    Private DefaultDenseVeldspar As Boolean = True
    Private DefaultIridescentGneiss As Boolean = True
    Private DefaultPrismaticGneiss As Boolean = True
    Private DefaultSilveryOmber As Boolean = True
    Private DefaultGoldenOmber As Boolean = True
    Private DefaultMagmaMercoxit As Boolean = True
    Private DefaultVitreousMercoxit As Boolean = True

    ' Default Shopping List Settings
    Private DefaultAlwaysonTop As Boolean = False
    Private DefaultUpdateAssetsWhenUsed As Boolean = False
    Private DefaultFees As Boolean = True
    Private DefaultCalcBuyBuyOrder As Boolean = True
    Private DefaultUsage As Boolean = True
    Private DefaultTotalItemTax As Boolean = True
    Private DefaultTotalItemBrokerFees As Boolean = True

    ' Assets - Item Checks
    Private DefaultAssetItemChecks As Boolean = True
    Private DefaultAssetItemTextFilter As String = ""
    Private DefaultAllItems As Boolean = True
    ' Assets - Main window 
    Private DefaultAssetType As String = "Both"
    Private DefaultAssetSortbyName As Boolean = True

    ' Local versions of settings
    Private ApplicationSettings As ApplicationSettings
    Private POSSettings As PlayerOwnedStationSettings
    Private BPSettings As BPTabSettings
    Private ManufacturingSettings As ManufacturingTabSettings
    Private DatacoreSettings As DataCoreTabSettings
    Private ReactionSettings As ReactionsTabSettings
    Private MiningSettings As MiningTabSettings
    Private UpdatePricesSettings As UpdatePriceTabSettings
    Private IndustryJobsColumnSettings As IndustryJobsColumnSettings
    Private ManufacturingTabColumnSettings As ManufacturingTabColumnSettings
    Private IndustryFlipBeltsSettings As IndustryFlipBeltSettings
    Private ShoppingListTabSettings As ShoppingListSettings
    Private IndustryTeamSettings As TeamSettings

    ' Facilities
    Private ManufacturingFacilitySettings As FacilitySettings
    Private ComponentsManufacturingFacilitySettings As FacilitySettings
    Private CapitalComponentsManufacturingFacilitySettings As FacilitySettings
    Private CapitalManufacturingFacilitySettings As FacilitySettings
    Private SuperManufacturingFacilitySettings As FacilitySettings
    Private T3CruiserManufacturingFacilitySettings As FacilitySettings
    Private T3DestroyerManufacturingFacilitySettings As FacilitySettings
    Private SubsystemManufacturingFacilitySettings As FacilitySettings
    Private BoosterManufacturingFacilitySettings As FacilitySettings
    Private CopyFacilitySettings As FacilitySettings
    Private InventionFacilitySettings As FacilitySettings
    Private T3InventionFacilitySettings As FacilitySettings
    Private NoPOSFacilitySettings As FacilitySettings
    Private POSFuelBlockFacilitySettings As FacilitySettings
    Private POSModuleFacilitySettings As FacilitySettings
    Private POSLargeShipFacilitySettings As FacilitySettings

    ' Teams
    Private BPManufacturingTeamSettings As TeamSettings
    Private BPComponentManufacturingTeamSettings As TeamSettings
    Private BPCopyTeamSettings As TeamSettings
    Private BPInventionTeamSettings As TeamSettings

    Private CalcManufacturingTeamSettings As TeamSettings
    Private CalcComponentManufacturingTeamSettings As TeamSettings
    Private CalcCopyTeamSettings As TeamSettings
    Private CalcInventionTeamSettings As TeamSettings

    ' Multiple versions of Asset windows
    Private AssetWindowSettingsDefault As AssetWindowSettings
    Private AssetWindowSettingsShoppingList As AssetWindowSettings

    ' 5 belt types
    Private IndustryBeltOreChecksSettings1 As IndustryBeltOreChecks
    Private IndustryBeltOreChecksSettings2 As IndustryBeltOreChecks
    Private IndustryBeltOreChecksSettings3 As IndustryBeltOreChecks
    Private IndustryBeltOreChecksSettings4 As IndustryBeltOreChecks
    Private IndustryBeltOreChecksSettings5 As IndustryBeltOreChecks

    Private Const AppSettingsFileName As String = "ApplicationSettings.xml"
    Private Const POSSettingsFileName As String = "POSSettings.xml"
    Private Const BPSettingsFileName As String = "BPTabSettings.xml"
    Private Const ManufacturingSettingsFileName As String = "ManufacturingTabSettings.xml"
    Private Const UpdatePricesFileName As String = "UpdatePricesSettings.xml"
    Private Const DatacoreSettingsFileName As String = "DatacoreSettings.xml"
    Private Const ReactionSettingsFileName As String = "ReactionTabSettings.xml"
    Private Const MiningSettingsFileName As String = "MiningTabSettings.xml"
    Private Const IndustryJobsColumnSettingsFileName As String = "IndustryJobsColumnSettings.xml"
    Private Const ManufacturingTabColumnSettingsFileName As String = "ManufacturingTabColumnSettings.xml"
    Private Const IndustryFlipBeltSettingsFileName As String = "IndustryFlipBeltSettings.xml"
    Private Const ShoppingListSettingsFileName As String = "ShoppingListSettings.xml"

    Private Const ManufacturingFacilitySettingsFileName As String = "ManufacturingFacilitySettings.xml"
    Private Const ComponentsManufacturingFacilitySettingsFileName As String = "ComponentsManufacturingFacilitySettings.xml"
    Private Const CapitalComponentsManufacturingFacilitySettingsFileName As String = "CapitalComponentsManufacturingFacilitySettings.xml"
    Private Const CapitalManufacturingFacilitySettingsFileName As String = "CapitalManufacturingFacilitySettings.xml"
    Private Const SuperCapitalManufacturingFacilitySettingsFileName As String = "SuperCapitalManufacturingFacilitySettings.xml"
    Private Const T3CruiserManufacturingFacilitySettingsFileName As String = "T3CruiserManufacturingFacilitySettings.xml"
    Private Const T3DestroyerManufacturingFacilitySettingsFileName As String = "T3DestroyerManufacturingFacilitySettings.xml"
    Private Const SubsystemManufacturingFacilitySettingsFileName As String = "SubsystemManufacturingFacilitySettings.xml"
    Private Const BoosterManufacturingFacilitySettingsFileName As String = "BoosterManufacturingFacilitySettings.xml"
    Private Const CopyFacilitySettingsFileName As String = "CopyFacilitySettings.xml"
    Private Const InventionFacilitySettingsFileName As String = "InventionFacilitySettings.xml"
    Private Const T3InventionFacilitySettingsFileName As String = "T3InventionFacilitySettings.xml"
    Private Const NoPoSFacilitySettingsFileName As String = "NoPOSFacilitySettings.xml"

    Private Const POSFuelBlockFacilitySettingsFileName As String = "POSFuelBlockFacilitySettings.xml"
    Private Const POSLargeShipFacilitySettingsFileName As String = "POSLargeShipFacilitySettings.xml"
    Private Const POSModuleFacilitySettingsFileName As String = "POSModuleFacilitySettings.xml"

    Private Const ManufacturingTeamSettingsFileName As String = "ManufacturingTeamSettings.xml"
    Private Const ComponentManufacturingTeamSettingsFileName As String = "ComponentsManufacturingTeamSettings.xml"
    Private Const CopyTeamSettingsFileName As String = "CopyTeamSettings.xml"
    Private Const InventionTeamSettingsFileName As String = "InventionTeamSettings.xml"

    ' 5 belts
    Private Const IndustryBeltOreChecksFileName1 As String = "IndustryBeltOreChecksSettings1.xml"
    Private Const IndustryBeltOreChecksFileName2 As String = "IndustryBeltOreChecksSettings2.xml"
    Private Const IndustryBeltOreChecksFileName3 As String = "IndustryBeltOreChecksSettings3.xml"
    Private Const IndustryBeltOreChecksFileName4 As String = "IndustryBeltOreChecksSettings4.xml"
    Private Const IndustryBeltOreChecksFileName5 As String = "IndustryBeltOreChecksSettings5.xml"

    ' Multiple asset windows
    Private Const AssetWindowFileNameDefault As String = "AssetWindowSettingsDefault.xml"
    Private Const AssetWindowFileNameShoppingList As String = "AssetWindowSettingsShoppingList.xml"

    Public Const SettingsFolder As String = "Settings/"

    Public Sub New()
        ApplicationSettings = Nothing
        MiningSettings = Nothing
        POSSettings = Nothing
        BPSettings = Nothing
        ManufacturingSettings = Nothing
        DatacoreSettings = Nothing
        ReactionSettings = Nothing
        MiningSettings = Nothing
        UpdatePricesSettings = Nothing
        IndustryJobsColumnSettings = Nothing
    End Sub

    ' Writes the sent settings to the sent file name
    Private Sub WriteSettingsToFile(FileName As String, Settings As Setting(), RootName As String)
        Dim i As Integer

        ' Create XmlWriterSettings.
        Dim XMLSettings As XmlWriterSettings = New XmlWriterSettings()
        XMLSettings.Indent = True

        If Not Directory.Exists(SettingsFolder) Then
            ' Create the settings folder
            Directory.CreateDirectory(SettingsFolder)
        End If

        ' Delete and make a fresh copy
        If File.Exists(SettingsFolder & FileName) Then
            File.Delete(SettingsFolder & FileName)
        End If

        ' Loop through the settings sent and output each name and value
        Using writer As XmlWriter = XmlWriter.Create(SettingsFolder & FileName, XMLSettings)
            writer.WriteStartDocument()
            writer.WriteStartElement(RootName) ' Root.

            ' Main loop
            For i = 0 To Settings.Count - 1
                writer.WriteElementString(Settings(i).Name, Settings(i).Value)
            Next

            ' End document.
            writer.WriteEndDocument()
        End Using

    End Sub

    ' Gets a value from a referenced XML file by searching for it
    Private Function GetSettingValue(ByRef XMLFile As String, ObjectType As SettingTypes, RootElement As String, ElementString As String, DefaultValue As Object) As Object
        Dim m_xmld As New XmlDocument
        Dim m_nodelist As XmlNodeList

        Dim TempValue As String

        'Load the Xml file
        m_xmld.Load(SettingsFolder & XMLFile)

        'Get the settings

        ' Get the cache update
        m_nodelist = m_xmld.SelectNodes("/" & RootElement & "/" & ElementString)

        If Not IsNothing(m_nodelist.Item(0)) Then
            ' Should only be one
            TempValue = m_nodelist.Item(0).InnerText

            ' If blank, then return default
            If TempValue = "" Then
                Return DefaultValue
            End If

            ' Found it, return the cast
            Select Case ObjectType
                Case SettingTypes.TypeBoolean
                    Return CBool(TempValue)
                Case SettingTypes.TypeDouble
                    Return CDbl(TempValue)
                Case SettingTypes.TypeInteger
                    Return CInt(TempValue)
                Case SettingTypes.TypeString
                    Return CStr(TempValue)
                Case SettingTypes.TypeLong
                    Return CLng(TempValue)
            End Select

        Else
            ' Doesn't exist, use default
            Return DefaultValue
        End If

        Return Nothing

    End Function

    Private Structure Setting
        Dim Name As String
        Dim Value As String

        Public Sub New(inName As String, inValue As String)
            Name = inName
            Value = inValue
        End Sub

    End Structure

    Private Enum SettingTypes
        TypeInteger = 1
        TypeDouble = 2
        TypeString = 3
        TypeBoolean = 4
        TypeLong = 5
    End Enum

#Region "Application Settings"

    ' Loads the settings for the user from the DB (for now) for the whole program
    Public Function LoadApplicationSettings() As ApplicationSettings
        Dim TempSettings As ApplicationSettings = Nothing

        Try
            If File.Exists(SettingsFolder & AppSettingsFileName) Then

                'Get the settings
                With TempSettings
                    .CheckforUpdatesonStart = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "CheckforUpdatesonStart", DefaultCheckUpdatesOnStart))
                    .LoadAssetsonStartup = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "LoadAssetsonStartup", DefaultLoadAssetsonStartup))
                    .LoadBPsonStartup = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "LoadbpsonStartup", DefaultLoadBPsonStartup))
                    .LoadCRESTTeamDataonStartup = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "LoadCRESTTeamDataonStartup", DefaultRefreshTeamCRESTDataonStartup))
                    .LoadCRESTMarketDataonStartup = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "LoadCRESTMarketDataonStartup", DefaultRefreshMarketCRESTDataonStartup))
                    .LoadCRESTFacilityDataonStartup = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "LoadCRESTFacilityDataonStartup", DefaultRefreshFacilityCRESTDataonStartup))
                    .DataExportFormat = CStr(GetSettingValue(AppSettingsFileName, SettingTypes.TypeString, "ApplicationSettings", "DataExportFormat", DefaultDataExportFormat))
                    .AllowSkillOverride = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "AllowSkillOverride", DefaultAllowSkillOverride))
                    .ShowToolTips = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "ShowToolTips", DefaultShowToolTips))
                    .RefiningImplantValue = CDbl(GetSettingValue(AppSettingsFileName, SettingTypes.TypeDouble, "ApplicationSettings", "RefiningImplantValue", DefaultImplantValues))
                    .ManufacturingImplantValue = CDbl(GetSettingValue(AppSettingsFileName, SettingTypes.TypeDouble, "ApplicationSettings", "ManufacturingImplantValue", DefaultImplantValues))
                    .CopyImplantValue = CDbl(GetSettingValue(AppSettingsFileName, SettingTypes.TypeDouble, "ApplicationSettings", "CopyImplantValue", DefaultImplantValues))
                    .BrokerCorpStanding = CDbl(GetSettingValue(AppSettingsFileName, SettingTypes.TypeDouble, "ApplicationSettings", "BrokerCorpStanding", DefaultBrokerCorpStanding))
                    .RefineCorpStanding = CDbl(GetSettingValue(AppSettingsFileName, SettingTypes.TypeDouble, "ApplicationSettings", "RefineCorpStanding", DefaultRefineCorpStanding))
                    .BrokerFactionStanding = CDbl(GetSettingValue(AppSettingsFileName, SettingTypes.TypeDouble, "ApplicationSettings", "BrokerFactionStanding", DefaultBrokerFactionStanding))
                    .DefaultBPME = CInt(GetSettingValue(AppSettingsFileName, SettingTypes.TypeInteger, "ApplicationSettings", "DefaultBPME", DefaultSettingME))
                    .DefaultBPTE = CInt(GetSettingValue(AppSettingsFileName, SettingTypes.TypeInteger, "ApplicationSettings", "DefaultBPTE", DefaultSettingTE))
                    .CheckBuildBuy = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "CheckBuildBuy", DefaultCheckBuildBuy))
                    .DisableSVR = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "DisableSVR", DefaultDisableSVR))
                    .RefiningEfficiency = CDbl(GetSettingValue(AppSettingsFileName, SettingTypes.TypeDouble, "ApplicationSettings", "RefiningEfficiency", DefaultRefiningEfficency))
                    .RefiningTax = CDbl(GetSettingValue(AppSettingsFileName, SettingTypes.TypeDouble, "ApplicationSettings", "RefiningTax", DefaultRefineTax))
                    .ShopListIncludeInventMats = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "ShopListIncludeInventMats", DefaultShopListIncludeInventMats))
                    .ShopListIncludeREMats = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "ShopListIncludeREMats", DefaultShopListIncludeREMats))
                    .SuggestBuildBPNotOwned = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "SuggestBuildBPNotOwned", DefaultSuggestBuildBPNotOwned))
                    .EVECentralRefreshInterval = CInt(GetSettingValue(AppSettingsFileName, SettingTypes.TypeInteger, "ApplicationSettings", "EVECentralRefreshInterval", DefaultEVECentralRefreshInterval))
                    .DisableSound = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "DisableSound", DefaultDisableSound))
                    .LinkBPTabtoFacilitySystem = CBool(GetSettingValue(AppSettingsFileName, SettingTypes.TypeBoolean, "ApplicationSettings", "LinkBPTabtoFacilitySystem", DefaultLinkTeamstoFacilitySystems))
                End With

                Select Case TempSettings.RefiningEfficiency
                    Case 0.5, 0.52, 0.53, 0.54, 0.57, 0.6
                        ' Do nothing
                    Case Else
                        ' Set to the default
                        TempSettings.RefiningEfficiency = DefaultRefiningEfficency
                End Select

            Else
                ' Load defaults 
                TempSettings = SetDefaultApplicationSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Application Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Some other error occured Load defaults 
            TempSettings = SetDefaultApplicationSettings()
        End Try

        ' Save them locally and then export
        ApplicationSettings = TempSettings

        Return TempSettings

    End Function

    ' Loads the defaults
    Public Function SetDefaultApplicationSettings() As ApplicationSettings
        Dim TempSettings As ApplicationSettings

        ' Load default settings
        TempSettings.CheckforUpdatesonStart = DefaultCheckUpdatesOnStart
        TempSettings.DataExportFormat = DefaultDataExportFormat
        TempSettings.ShowToolTips = DefaultShowToolTips
        TempSettings.LoadAssetsonStartup = DefaultLoadAssetsonStartup
        TempSettings.LoadBPsonStartup = DefaultLoadBPsonStartup
        TempSettings.LoadCRESTTeamDataonStartup = DefaultRefreshTeamCRESTDataonStartup
        TempSettings.LoadCRESTMarketDataonStartup = DefaultRefreshMarketCRESTDataonStartup
        TempSettings.LoadCRESTFacilityDataonStartup = DefaultRefreshFacilityCRESTDataonStartup
        TempSettings.DisableSound = DefaultDisableSound
        TempSettings.ManufacturingImplantValue = DefaultImplantValues
        TempSettings.RefiningImplantValue = DefaultImplantValues
        TempSettings.CopyImplantValue = DefaultImplantValues

        ' Station Standings for building and selling
        TempSettings.BrokerCorpStanding = DefaultBrokerCorpStanding
        TempSettings.BrokerFactionStanding = DefaultBrokerFactionStanding
        TempSettings.RefineCorpStanding = DefaultRefineCorpStanding

        TempSettings.CheckBuildBuy = DefaultCheckBuildBuy

        TempSettings.DefaultBPME = DefaultSettingME
        TempSettings.DefaultBPTE = DefaultSettingTE

        TempSettings.RefiningEfficiency = DefaultRefiningEfficency

        TempSettings.RefiningTax = DefaultRefineTax

        TempSettings.DisableSVR = DefaultDisableSVR
        TempSettings.SuggestBuildBPNotOwned = DefaultSuggestBuildBPNotOwned
        TempSettings.LinkBPTabtoFacilitySystem = DefaultLinkTeamstoFacilitySystems

        TempSettings.ShopListIncludeInventMats = DefaultShopListIncludeInventMats
        TempSettings.ShopListIncludeREMats = DefaultShopListIncludeREMats

        TempSettings.EVECentralRefreshInterval = DefaultEVECentralRefreshInterval

        ' Save locally
        ApplicationSettings = TempSettings
        Return TempSettings

    End Function

    ' Saves the application settings to XML
    Public Sub SaveApplicationSettings(SentSettings As ApplicationSettings)
        Dim ApplicationSettingsList(25) As Setting

        Try
            ApplicationSettingsList(0) = New Setting("CheckforUpdatesonStart", CStr(SentSettings.CheckforUpdatesonStart))
            ApplicationSettingsList(1) = New Setting("DataExportFormat", CStr(SentSettings.DataExportFormat))
            ApplicationSettingsList(2) = New Setting("AllowSkillOverride", CStr(SentSettings.AllowSkillOverride))
            ApplicationSettingsList(3) = New Setting("ShowToolTips", CStr(SentSettings.ShowToolTips))
            ApplicationSettingsList(4) = New Setting("RefiningImplantValue", CStr(SentSettings.RefiningImplantValue))
            ApplicationSettingsList(5) = New Setting("ManufacturingImplantValue", CStr(SentSettings.ManufacturingImplantValue))
            ApplicationSettingsList(6) = New Setting("CopyImplantValue", CStr(SentSettings.CopyImplantValue))
            ApplicationSettingsList(7) = New Setting("BrokerCorpStanding", CStr(SentSettings.BrokerCorpStanding))
            ApplicationSettingsList(8) = New Setting("RefineCorpStanding", CStr(SentSettings.RefineCorpStanding))
            ApplicationSettingsList(9) = New Setting("BrokerFactionStanding", CStr(SentSettings.BrokerFactionStanding))
            ApplicationSettingsList(10) = New Setting("DefaultBPME", CStr(SentSettings.DefaultBPME))
            ApplicationSettingsList(11) = New Setting("DefaultBPTE", CStr(SentSettings.DefaultBPTE))
            ApplicationSettingsList(12) = New Setting("CheckBuildBuy", CStr(SentSettings.CheckBuildBuy))
            ApplicationSettingsList(13) = New Setting("RefiningEfficiency", CStr(SentSettings.RefiningEfficiency))
            ApplicationSettingsList(14) = New Setting("RefiningTax", CStr(SentSettings.RefiningTax))
            ApplicationSettingsList(15) = New Setting("ShopListIncludeInventMats", CStr(SentSettings.ShopListIncludeInventMats))
            ApplicationSettingsList(16) = New Setting("ShopListIncludeREMats", CStr(SentSettings.ShopListIncludeREMats))
            ApplicationSettingsList(17) = New Setting("SuggestBuildBPNotOwned", CStr(SentSettings.SuggestBuildBPNotOwned))
            ApplicationSettingsList(18) = New Setting("EVECentralRefreshInterval", CStr(SentSettings.EVECentralRefreshInterval))
            ApplicationSettingsList(19) = New Setting("LoadAssetsonStartup", CStr(SentSettings.LoadAssetsonStartup))
            ApplicationSettingsList(20) = New Setting("DisableSound", CStr(SentSettings.DisableSound))
            ApplicationSettingsList(21) = New Setting("LoadbpsonStartup", CStr(SentSettings.LoadBPsonStartup))
            ApplicationSettingsList(22) = New Setting("LoadCRESTTeamDataonStartup", CStr(SentSettings.LoadCRESTTeamDataonStartup))
            ApplicationSettingsList(23) = New Setting("LoadCRESTFacilityDataonStartup", CStr(SentSettings.LoadCRESTFacilityDataonStartup))
            ApplicationSettingsList(24) = New Setting("LoadCRESTMarketDataonStartup", CStr(SentSettings.LoadCRESTMarketDataonStartup))
            ApplicationSettingsList(25) = New Setting("LinkBPTabtoFacilitySystem", CStr(SentSettings.LinkBPTabtoFacilitySystem))

            Call WriteSettingsToFile(AppSettingsFileName, ApplicationSettingsList, "ApplicationSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Application Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the application settings
    Public Function GetApplicationSettings() As ApplicationSettings
        Return ApplicationSettings
    End Function

#End Region

#Region "Shopping List Settings"

    ' Loads the POS tower settings from XML setting file
    Public Function LoadShoppingListSettings() As ShoppingListSettings
        Dim TempSettings As ShoppingListSettings = Nothing

        Try
            If File.Exists(SettingsFolder & ShoppingListSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .DataExportFormat = CStr(GetSettingValue(ShoppingListSettingsFileName, SettingTypes.TypeString, "ShoppingListSettings", "DataExportFormat", DefaultDataExportFormat))
                    .AlwaysonTop = CBool(GetSettingValue(ShoppingListSettingsFileName, SettingTypes.TypeBoolean, "ShoppingListSettings", "AlwaysonTop", DefaultAlwaysonTop))
                    .UpdateAssetsWhenUsed = CBool(GetSettingValue(ShoppingListSettingsFileName, SettingTypes.TypeBoolean, "ShoppingListSettings", "UpdateAssetsWhenUsed", DefaultUpdateAssetsWhenUsed))
                    .Fees = CBool(GetSettingValue(ShoppingListSettingsFileName, SettingTypes.TypeBoolean, "ShoppingListSettings", "Fees", DefaultFees))
                    .CalcBuyBuyOrder = CBool(GetSettingValue(ShoppingListSettingsFileName, SettingTypes.TypeBoolean, "ShoppingListSettings", "CalcBuyBuyOrder", DefaultCalcBuyBuyOrder))
                    .Usage = CBool(GetSettingValue(ShoppingListSettingsFileName, SettingTypes.TypeBoolean, "ShoppingListSettings", "Usage", DefaultUsage))
                    .TotalItemTax = CBool(GetSettingValue(ShoppingListSettingsFileName, SettingTypes.TypeBoolean, "ShoppingListSettings", "TotalItemTax", DefaultTotalItemTax))
                    .TotalItemBrokerFees = CBool(GetSettingValue(ShoppingListSettingsFileName, SettingTypes.TypeBoolean, "ShoppingListSettings", "TotalItemBrokerFees", DefaultTotalItemBrokerFees))
                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultShopingListSettings()
            End If
        Catch ex As Exception
            MsgBox("An error occured when loading Shopping List Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultShopingListSettings()
        End Try

        ' Save them locally and then export
        ShoppingListTabSettings = TempSettings

        Return TempSettings

    End Function

    ' Load defaults 
    Public Function SetDefaultShopingListSettings() As ShoppingListSettings
        Dim TempSettings As ShoppingListSettings = Nothing

        ' Load defaults 
        TempSettings.DataExportFormat = DefaultDataExportFormat
        TempSettings.AlwaysonTop = DefaultAlwaysonTop
        TempSettings.UpdateAssetsWhenUsed = DefaultUpdateAssetsWhenUsed
        TempSettings.UpdateAssetsWhenUsed = DefaultUpdateAssetsWhenUsed
        TempSettings.Fees = DefaultFees
        TempSettings.CalcBuyBuyOrder = DefaultCalcBuyBuyOrder
        TempSettings.Usage = DefaultUsage
        TempSettings.TotalItemTax = DefaultTotalItemTax
        TempSettings.TotalItemBrokerFees = DefaultTotalItemBrokerFees

        ShoppingListTabSettings = TempSettings

        Return TempSettings

    End Function

    ' Saves the Shopping List Settings to XML
    Public Sub SaveShoppingListSettings(SentSettings As ShoppingListSettings)
        Dim ShoppingListSettingsList(7) As Setting

        Try
            ShoppingListSettingsList(0) = New Setting("DataExportFormat", CStr(SentSettings.DataExportFormat))
            ShoppingListSettingsList(1) = New Setting("AlwaysonTop", CStr(SentSettings.AlwaysonTop))
            ShoppingListSettingsList(2) = New Setting("UpdateAssetsWhenUsed", CStr(SentSettings.UpdateAssetsWhenUsed))
            ShoppingListSettingsList(3) = New Setting("Fees", CStr(SentSettings.Fees))
            ShoppingListSettingsList(4) = New Setting("CalcBuyBuyOrder", CStr(SentSettings.CalcBuyBuyOrder))
            ShoppingListSettingsList(5) = New Setting("Usage", CStr(SentSettings.Usage))
            ShoppingListSettingsList(6) = New Setting("TotalItemTax", CStr(SentSettings.TotalItemTax))
            ShoppingListSettingsList(7) = New Setting("TotalItemBrokerFees", CStr(SentSettings.TotalItemBrokerFees))

            Call WriteSettingsToFile(ShoppingListSettingsFileName, ShoppingListSettingsList, "ShoppingListSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Shopping List Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the Shopping List Settings
    Public Function GetShoppingListSettings() As ShoppingListSettings
        Return ShoppingListTabSettings
    End Function

#End Region

#Region "BP Tab Settings"

    ' Loads the tab settings
    Public Function LoadBPSettings() As BPTabSettings
        Dim TempSettings As BPTabSettings = Nothing

        Try
            If File.Exists(SettingsFolder & BPSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .BlueprintTypeSelection = CStr(GetSettingValue(BPSettingsFileName, SettingTypes.TypeString, "BPSettings", "BlueprintTypeSelection", DefaultBPSelectionType))
                    .Tech1Check = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "Tech1Check", DefaultBPTechChecks))
                    .Tech2Check = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "Tech2Check", DefaultBPTechChecks))
                    .Tech3Check = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "Tech3Check", DefaultBPTechChecks))
                    .TechStorylineCheck = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "TechStorylineCheck", DefaultBPTechChecks))
                    .TechFactionCheck = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "TechFactionCheck", DefaultBPTechChecks))
                    .TechPirateCheck = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "TechPirateCheck", DefaultBPTechChecks))
                    .IncludeUsage = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeUsage", DefaultBPIncludeUsage))
                    .IncludeTaxes = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeTaxes", DefaultBPIncludeTaxes))
                    .PricePerUnit = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "PricePerUnit", DefaultBPPricePerUnit))
                    .IncludeInventionCost = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeInventionCost", DefaultBPIncludeInventionCost))
                    .IncludeInventionTime = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeInventionTime", DefaultBPIncludeInventionTime))
                    .IncludeCopyCost = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeCopyCost", DefaultBPIncludecopyCost))
                    .IncludeCopyTime = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeCopyTime", DefaultBPIncludeCopyTime))
                    .IncludeT3Cost = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeT3Cost", DefaultBPIncludeT3Cost))
                    .IncludeT3Time = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeT3Time", DefaultBPIncludeT3Time))
                    .ProductionLines = CInt(GetSettingValue(BPSettingsFileName, SettingTypes.TypeInteger, "BPSettings", "ProductionLines", DefaultBPProductionLines))
                    .LaboratoryLines = CInt(GetSettingValue(BPSettingsFileName, SettingTypes.TypeInteger, "BPSettings", "LaboratoryLines", DefaultBPLaboratoryLines))
                    .T3Lines = CInt(GetSettingValue(BPSettingsFileName, SettingTypes.TypeInteger, "BPSettings", "RELines", DefaultBPRELines))
                    .SmallCheck = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "SmallCheck", DefaultSizeChecks))
                    .MediumCheck = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "SmallCheck", DefaultSizeChecks))
                    .LargeCheck = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "SmallCheck", DefaultSizeChecks))
                    .XLCheck = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "SmallCheck", DefaultSizeChecks))
                    .IncludeFees = CBool(GetSettingValue(BPSettingsFileName, SettingTypes.TypeBoolean, "BPSettings", "IncludeFees", DefaultBPIncludeFees))
                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultBPSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Application Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultBPSettings()
        End Try

        ' Save them locally and then export
        BPSettings = TempSettings

        Return TempSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveBPSettings(SentSettings As BPTabSettings)
        Dim BPSettingsList(23) As Setting

        Try
            BPSettingsList(0) = New Setting("BlueprintTypeSelection", CStr(SentSettings.BlueprintTypeSelection))
            BPSettingsList(1) = New Setting("Tech1Check", CStr(SentSettings.Tech1Check))
            BPSettingsList(2) = New Setting("Tech2Check", CStr(SentSettings.Tech2Check))
            BPSettingsList(3) = New Setting("Tech3Check", CStr(SentSettings.Tech3Check))
            BPSettingsList(4) = New Setting("TechStorylineCheck", CStr(SentSettings.TechStorylineCheck))
            BPSettingsList(5) = New Setting("TechFactionCheck", CStr(SentSettings.TechFactionCheck))
            BPSettingsList(6) = New Setting("TechPirateCheck", CStr(SentSettings.TechPirateCheck))
            BPSettingsList(7) = New Setting("IncludeUsage", CStr(SentSettings.IncludeUsage))
            BPSettingsList(8) = New Setting("IncludeTaxes", CStr(SentSettings.IncludeTaxes))
            BPSettingsList(9) = New Setting("PricePerUnit", CStr(SentSettings.PricePerUnit))
            BPSettingsList(10) = New Setting("ProductionLines", CStr(SentSettings.ProductionLines))
            BPSettingsList(11) = New Setting("LaboratoryLines", CStr(SentSettings.LaboratoryLines))
            BPSettingsList(12) = New Setting("RELines", CStr(SentSettings.T3Lines))
            BPSettingsList(13) = New Setting("SmallCheck", CStr(SentSettings.SmallCheck))
            BPSettingsList(14) = New Setting("MediumCheck", CStr(SentSettings.MediumCheck))
            BPSettingsList(15) = New Setting("LargeCheck", CStr(SentSettings.LargeCheck))
            BPSettingsList(16) = New Setting("XLCheck", CStr(SentSettings.XLCheck))
            BPSettingsList(17) = New Setting("IncludeFees", CStr(SentSettings.IncludeFees))

            BPSettingsList(18) = New Setting("IncludeInventionCost", CStr(SentSettings.IncludeInventionCost))
            BPSettingsList(19) = New Setting("IncludeInventionTime", CStr(SentSettings.IncludeInventionTime))
            BPSettingsList(20) = New Setting("IncludeCopyCost", CStr(SentSettings.IncludeCopyCost))
            BPSettingsList(21) = New Setting("IncludeCopyTime", CStr(SentSettings.IncludeCopyTime))
            BPSettingsList(22) = New Setting("IncludeT3Cost", CStr(SentSettings.IncludeT3Cost))
            BPSettingsList(23) = New Setting("IncludeT3Time", CStr(SentSettings.IncludeT3Time))

            Call WriteSettingsToFile(BPSettingsFileName, BPSettingsList, "BPSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving BP Tab Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetBPSettings() As BPTabSettings
        Return BPSettings
    End Function

    ' Loads the defaults
    Public Function SetDefaultBPSettings() As BPTabSettings
        Dim LocalSettings As BPTabSettings

        LocalSettings.BlueprintTypeSelection = DefaultBPSelectionType
        LocalSettings.Tech1Check = DefaultBPTechChecks
        LocalSettings.Tech2Check = DefaultBPTechChecks
        LocalSettings.Tech3Check = DefaultBPTechChecks
        LocalSettings.TechStorylineCheck = DefaultBPTechChecks
        LocalSettings.TechFactionCheck = DefaultBPTechChecks
        LocalSettings.TechPirateCheck = DefaultBPTechChecks
        LocalSettings.IncludeUsage = DefaultBPIncludeFees
        LocalSettings.IncludeTaxes = DefaultBPIncludeTaxes
        LocalSettings.IncludeFees = DefaultIncludeBrokerFees
        LocalSettings.PricePerUnit = DefaultBPPricePerUnit
        LocalSettings.ProductionLines = DefaultBPProductionLines
        LocalSettings.LaboratoryLines = DefaultBPLaboratoryLines
        LocalSettings.T3Lines = DefaultBPRELines
        LocalSettings.SmallCheck = DefaultSizeChecks
        LocalSettings.MediumCheck = DefaultSizeChecks
        LocalSettings.LargeCheck = DefaultSizeChecks
        LocalSettings.XLCheck = DefaultSizeChecks

        LocalSettings.IncludeInventionCost = DefaultBPIncludeInventionCost
        LocalSettings.IncludeInventionTime = DefaultBPIncludeInventionTime
        LocalSettings.IncludeCopyCost = DefaultBPIncludecopyCost
        LocalSettings.IncludeCopyTime = DefaultBPIncludeCopyTime
        LocalSettings.IncludeT3Cost = DefaultBPIncludeT3Cost
        LocalSettings.IncludeT3Time = DefaultBPIncludeT3Time

        ' Save locally
        BPSettings = LocalSettings
        Return LocalSettings

    End Function

#End Region

#Region "Update Price Tab Settings"

    ' Loads the tab settings
    Public Function LoadUpdatePricesSettings() As UpdatePriceTabSettings
        Dim TempSettings As UpdatePriceTabSettings = Nothing

        Try
            If File.Exists(SettingsFolder & UpdatePricesFileName) Then

                'Get the settings
                With TempSettings
                    .AllRawMats = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "AllRawMats", DefaultPriceChecks))
                    .Minerals = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Minerals", DefaultPriceChecks))
                    .IceProducts = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "IceProducts", DefaultPriceChecks))
                    .Gas = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Gas", DefaultPriceChecks))
                    .Misc = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Misc", DefaultPriceChecks))
                    .AncientRelics = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "AncientRelics", DefaultPriceChecks))
                    .AncientSalvage = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "AncientSalvage", DefaultPriceChecks))
                    .Salvage = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Salvage", DefaultPriceChecks))
                    .StationComponents = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "StationComponents", DefaultPriceChecks))
                    .Planetary = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Planetary", DefaultPriceChecks))
                    .Datacores = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Datacores", DefaultPriceChecks))
                    .Decryptors = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Decryptors", DefaultPriceChecks))
                    .Deployables = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Deployables", DefaultPriceChecks))
                    .Celestials = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Celestials", DefaultPriceChecks))
                    .Deployables = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Deployables", DefaultPriceChecks))
                    .Implants = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Implants", DefaultPriceChecks))
                    .RawMats = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "RawMats", DefaultPriceChecks))
                    .ProcessedMats = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "ProcessedMats", DefaultPriceChecks))
                    .AdvancedMats = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "AdvancedMats", DefaultPriceChecks))
                    .MatsandCompounds = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "MatsandCompounds", DefaultPriceChecks))
                    .DroneComponents = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "DroneComponents", DefaultPriceChecks))
                    .BoosterMats = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "BoosterMats", DefaultPriceChecks))
                    .Polymers = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Polymers", DefaultPriceChecks))
                    .Asteroids = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Asteroids", DefaultPriceChecks))
                    .AllManufacturedItems = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "AllManufacturedItems", DefaultPriceChecks))
                    .Ships = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Ships", DefaultPriceChecks))
                    .Modules = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Modules", DefaultPriceChecks))
                    .Drones = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Drones", DefaultPriceChecks))
                    .Boosters = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Boosters", DefaultPriceChecks))
                    .Rigs = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Rigs", DefaultPriceChecks))
                    .Charges = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Charges", DefaultPriceChecks))
                    .Subsystems = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Subsystems", DefaultPriceChecks))
                    .Structures = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Structures", DefaultPriceChecks))
                    .Tools = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Tools", DefaultPriceChecks))
                    .DataInterfaces = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "DataInterfaces", DefaultPriceChecks))
                    .CapT2Components = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "CapT2Components", DefaultPriceChecks))
                    .CapitalComponents = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "CapitalComponents", DefaultPriceChecks))
                    .Components = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Components", DefaultPriceChecks))
                    .Hybrid = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Hybrid", DefaultPriceChecks))
                    .FuelBlocks = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "FuelBlocks", DefaultPriceChecks))
                    .T1 = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "T1", DefaultPriceChecks))
                    .T2 = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "T2", DefaultPriceChecks))
                    .T3 = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "T3", DefaultPriceChecks))
                    .Faction = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Faction", DefaultPriceChecks))
                    .Pirate = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Pirate", DefaultPriceChecks))
                    .Storyline = CBool(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeBoolean, "UpdatePricesSettings", "Storyline", DefaultPriceChecks))

                    Dim TempRegions As String = CStr(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeString, "UpdatePricesSettings", "SelectedRegions", DefaultPriceRegion))
                    Dim RegionList As New List(Of String)
                    Dim RegionCount As Integer

                    If TempRegions <> "0" Then
                        RegionCount = System.Text.RegularExpressions.Regex.Matches(TempRegions, Regex.Escape(",")).Count + 1 ' Add one for last item + 1 ' Add one for last item
                    End If

                    Dim ReaderStartPosition As Integer = 0
                    Dim CommaLoc As Integer

                    For i = 0 To RegionCount - 1
                        CommaLoc = InStr(TempRegions.Substring(ReaderStartPosition), ",")
                        If CommaLoc <> 0 Then
                            RegionList.Add(TempRegions.Substring(ReaderStartPosition, CommaLoc - 1))
                        Else ' At the end
                            RegionList.Add(TempRegions.Substring(ReaderStartPosition))
                        End If
                        ReaderStartPosition = ReaderStartPosition + CommaLoc
                    Next

                    .SelectedRegions = RegionList
                    .SelectedSystem = CStr(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeString, "UpdatePricesSettings", "SelectedSystem", DefaultPriceSystem))
                    .PriceImportType = CStr(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeString, "UpdatePricesSettings", "PriceImportType", DefaultPriceImportPriceType))
                    .ItemsCombo = CStr(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeString, "UpdatePricesSettings", "ItemsCombo", DefaultPriceItemsCombo))
                    .RawMatsCombo = CStr(GetSettingValue(UpdatePricesFileName, SettingTypes.TypeString, "UpdatePricesSettings", "RawMatsCombo", DefaultPriceRawMatsCombo))

                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultUpdatePriceSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Update Prices Tab Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultUpdatePriceSettings()
        End Try

        ' Save them locally and then export
        UpdatePricesSettings = TempSettings

        Return TempSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveUpdatePricesSettings(PriceSettings As UpdatePriceTabSettings)
        Dim UpdatePricesSettingsList(49) As Setting

        Try
            UpdatePricesSettingsList(0) = New Setting("AllRawMats", CStr(PriceSettings.AllRawMats))
            UpdatePricesSettingsList(1) = New Setting("Minerals", CStr(PriceSettings.Minerals))
            UpdatePricesSettingsList(2) = New Setting("IceProducts", CStr(PriceSettings.IceProducts))
            UpdatePricesSettingsList(3) = New Setting("Gas", CStr(PriceSettings.Gas))
            UpdatePricesSettingsList(4) = New Setting("AncientRelics", CStr(PriceSettings.AncientRelics))
            UpdatePricesSettingsList(5) = New Setting("AncientSalvage", CStr(PriceSettings.AncientSalvage))
            UpdatePricesSettingsList(6) = New Setting("Salvage", CStr(PriceSettings.Salvage))
            UpdatePricesSettingsList(7) = New Setting("StationComponents", CStr(PriceSettings.StationComponents))
            UpdatePricesSettingsList(8) = New Setting("Planetary", CStr(PriceSettings.Planetary))
            UpdatePricesSettingsList(9) = New Setting("Datacores", CStr(PriceSettings.Datacores))
            UpdatePricesSettingsList(10) = New Setting("Decryptors", CStr(PriceSettings.Decryptors))
            UpdatePricesSettingsList(11) = New Setting("RawMats", CStr(PriceSettings.RawMats))
            UpdatePricesSettingsList(12) = New Setting("ProcessedMats", CStr(PriceSettings.ProcessedMats))
            UpdatePricesSettingsList(13) = New Setting("AdvancedMats", CStr(PriceSettings.AdvancedMats))
            UpdatePricesSettingsList(14) = New Setting("MatsandCompounds", CStr(PriceSettings.MatsandCompounds))
            UpdatePricesSettingsList(15) = New Setting("DroneComponents", CStr(PriceSettings.DroneComponents))
            UpdatePricesSettingsList(16) = New Setting("BoosterMats", CStr(PriceSettings.BoosterMats))
            UpdatePricesSettingsList(17) = New Setting("Polymers", CStr(PriceSettings.Polymers))
            UpdatePricesSettingsList(18) = New Setting("AllManufacturedItems", CStr(PriceSettings.AllManufacturedItems))
            UpdatePricesSettingsList(19) = New Setting("Ships", CStr(PriceSettings.Ships))
            UpdatePricesSettingsList(20) = New Setting("Modules", CStr(PriceSettings.Modules))
            UpdatePricesSettingsList(21) = New Setting("Drones", CStr(PriceSettings.Drones))
            UpdatePricesSettingsList(22) = New Setting("Boosters", CStr(PriceSettings.Boosters))
            UpdatePricesSettingsList(23) = New Setting("Rigs", CStr(PriceSettings.Rigs))
            UpdatePricesSettingsList(24) = New Setting("Charges", CStr(PriceSettings.Charges))
            UpdatePricesSettingsList(25) = New Setting("Subsystems", CStr(PriceSettings.Subsystems))
            UpdatePricesSettingsList(26) = New Setting("Structures", CStr(PriceSettings.Structures))
            UpdatePricesSettingsList(27) = New Setting("Tools", CStr(PriceSettings.Tools))
            UpdatePricesSettingsList(28) = New Setting("DataInterfaces", CStr(PriceSettings.DataInterfaces))
            UpdatePricesSettingsList(29) = New Setting("CapT2Components", CStr(PriceSettings.CapT2Components))
            UpdatePricesSettingsList(30) = New Setting("CapitalComponents", CStr(PriceSettings.CapitalComponents))
            UpdatePricesSettingsList(31) = New Setting("Components", CStr(PriceSettings.Components))
            UpdatePricesSettingsList(32) = New Setting("Hybrid", CStr(PriceSettings.Hybrid))
            UpdatePricesSettingsList(33) = New Setting("FuelBlocks", CStr(PriceSettings.FuelBlocks))
            UpdatePricesSettingsList(34) = New Setting("T1", CStr(PriceSettings.T1))
            UpdatePricesSettingsList(35) = New Setting("T2", CStr(PriceSettings.T2))
            UpdatePricesSettingsList(36) = New Setting("T3", CStr(PriceSettings.T3))
            UpdatePricesSettingsList(37) = New Setting("Faction", CStr(PriceSettings.Faction))
            UpdatePricesSettingsList(38) = New Setting("Pirate", CStr(PriceSettings.Pirate))
            UpdatePricesSettingsList(39) = New Setting("Storyline", CStr(PriceSettings.Storyline))
            Dim RegionList As String = ""
            If Not IsNothing(PriceSettings.SelectedRegions) Then
                For i = 0 To PriceSettings.SelectedRegions.Count - 1
                    RegionList = RegionList & PriceSettings.SelectedRegions(i) & ","
                Next
                If RegionList <> "" Then
                    ' Strip last comma
                    RegionList = RegionList.Substring(0, Len(RegionList) - 1)
                End If
            Else
                RegionList = "0"
            End If
            UpdatePricesSettingsList(40) = New Setting("SelectedRegions", RegionList)
            UpdatePricesSettingsList(41) = New Setting("SelectedSystem", CStr(PriceSettings.SelectedSystem))
            UpdatePricesSettingsList(42) = New Setting("PriceImportType", CStr(PriceSettings.PriceImportType))
            UpdatePricesSettingsList(43) = New Setting("ItemsCombo", CStr(PriceSettings.ItemsCombo))
            UpdatePricesSettingsList(44) = New Setting("RawMatsCombo", CStr(PriceSettings.RawMatsCombo))

            UpdatePricesSettingsList(45) = New Setting("Asteroids", CStr(PriceSettings.Asteroids))
            UpdatePricesSettingsList(46) = New Setting("Misc", CStr(PriceSettings.Misc))

            UpdatePricesSettingsList(47) = New Setting("Deployables", CStr(PriceSettings.Deployables))
            UpdatePricesSettingsList(48) = New Setting("Celestials", CStr(PriceSettings.Celestials))
            UpdatePricesSettingsList(49) = New Setting("Implants", CStr(PriceSettings.Implants))

            Call WriteSettingsToFile(UpdatePricesFileName, UpdatePricesSettingsList, "UpdatePricesSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Update Prices Tab Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetUpdatePricesSettings() As UpdatePriceTabSettings
        Return UpdatePricesSettings
    End Function

    Public Function SetDefaultUpdatePriceSettings() As UpdatePriceTabSettings
        Dim LocalSettings As UpdatePriceTabSettings

        With LocalSettings
            .AllRawMats = DefaultPriceChecks
            .Minerals = DefaultPriceChecks
            .IceProducts = DefaultPriceChecks
            .Gas = DefaultPriceChecks
            .Misc = DefaultPriceChecks
            .AncientRelics = DefaultPriceChecks
            .AncientSalvage = DefaultPriceChecks
            .Salvage = DefaultPriceChecks
            .StationComponents = DefaultPriceChecks
            .Planetary = DefaultPriceChecks
            .Datacores = DefaultPriceChecks
            .Decryptors = DefaultPriceChecks
            .RawMats = DefaultPriceChecks
            .ProcessedMats = DefaultPriceChecks
            .AdvancedMats = DefaultPriceChecks
            .MatsandCompounds = DefaultPriceChecks
            .DroneComponents = DefaultPriceChecks
            .BoosterMats = DefaultPriceChecks
            .Polymers = DefaultPriceChecks
            .Asteroids = DefaultPriceChecks
            .AllManufacturedItems = DefaultPriceChecks
            .Ships = DefaultPriceChecks
            .Modules = DefaultPriceChecks
            .Drones = DefaultPriceChecks
            .Boosters = DefaultPriceChecks
            .Rigs = DefaultPriceChecks
            .Charges = DefaultPriceChecks
            .Subsystems = DefaultPriceChecks
            .Structures = DefaultPriceChecks
            .Tools = DefaultPriceChecks
            .DataInterfaces = DefaultPriceChecks
            .CapT2Components = DefaultPriceChecks
            .CapitalComponents = DefaultPriceChecks
            .Components = DefaultPriceChecks
            .Hybrid = DefaultPriceChecks
            .FuelBlocks = DefaultPriceChecks
            .Implants = DefaultPriceChecks
            .Celestials = DefaultPriceChecks
            .Deployables = DefaultPriceChecks
            .T1 = DefaultPriceChecks
            .T2 = DefaultPriceChecks
            .T3 = DefaultPriceChecks
            .Faction = DefaultPriceChecks
            .Pirate = DefaultPriceChecks
            .Storyline = DefaultPriceChecks
            .SelectedRegions = Nothing
            .SelectedSystem = DefaultPriceSystem
            .PriceImportType = DefaultPriceImportPriceType
            .ItemsCombo = DefaultPriceItemsCombo
            .RawMatsCombo = DefaultPriceRawMatsCombo
        End With

        ' Save locally
        UpdatePricesSettings = LocalSettings
        Return LocalSettings

    End Function

#End Region

#Region "Manufacturing Tab Settings"

    ' Loads the tab settings
    Public Function LoadManufacturingSettings() As ManufacturingTabSettings
        Dim TempSettings As ManufacturingTabSettings = Nothing

        Try
            If File.Exists(SettingsFolder & ManufacturingSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .BlueprintType = CStr(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeString, "ManufacturingSettings", "BlueprintType", DefaultBlueprintType))
                    .CheckTech1 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckTech1", DefaultCheckTech1))
                    .CheckTech2 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckTech2", DefaultCheckTech2))
                    .CheckTech3 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckTech3", DefaultCheckTech3))
                    .CheckTechStoryline = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckTechStoryline", DefaultCheckTechStoryline))
                    .CheckTechNavy = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckTechNavy", DefaultCheckTechNavy))
                    .CheckTechPirate = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckTechPirate", DefaultCheckTechPirate))
                    .ItemTypeFilter = CStr(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeString, "ManufacturingSettings", "ItemTypeFilter", DefaultItemTypeFilter))
                    .TextItemFilter = CStr(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeString, "ManufacturingSettings", "TextItemFilter", DefaultTextItemFilter))
                    .CheckBPTypeShips = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeShips", DefaultCheckBPTypeShips))
                    .CheckBPTypeDrones = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeDrones", DefaultCheckBPTypeDrones))
                    .CheckBPTypeComponents = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeComponents", DefaultCheckBPTypeComponents))
                    .CheckBPTypeStructures = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeStructures", DefaultCheckBPTypeStructures))
                    .CheckBPTypeMisc = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeMisc", DefaultCheckBPTypeTools))
                    .CheckBPTypeModules = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeModules", DefaultCheckBPTypeModules))
                    .CheckBPTypeAmmoCharges = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeAmmoCharges", DefaultCheckBPTypeAmmoCharges))
                    .CheckBPTypeRigs = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeRigs", DefaultCheckBPTypeRigs))
                    .CheckBPTypeSubsystems = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeSubsystems", DefaultCheckBPTypeSubsystems))
                    .CheckBPTypeBoosters = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeBoosters", DefaultCheckBPTypeBoosters))
                    .CheckBPTypeDeployables = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeDeployables", DefaultCheckBPTypeDeployables))
                    .CheckBPTypeCelestials = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeCelestials", DefaultCheckBPTypeCelestials))
                    .CheckBPTypeStationParts = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckBPTypeStationParts", DefaultCheckBPTypeStationParts))
                    .AveragePriceDuration = CStr(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeString, "ManufacturingSettings", "AveragePriceDuration", DefaultAveragePriceDuration))
                    .CheckDecryptorNone = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptorNone", DefaultCheckDecryptorNone))
                    .CheckDecryptor06 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptor06", DefaultCheckDecryptor06))
                    .CheckDecryptor09 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptor09", DefaultCheckDecryptor09))
                    .CheckDecryptor10 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptor10", DefaultCheckDecryptor10))
                    .CheckDecryptor11 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptor11", DefaultCheckDecryptor11))
                    .CheckDecryptor12 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptor12", DefaultCheckDecryptor12))
                    .CheckDecryptor15 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptor15", DefaultCheckDecryptor15))
                    .CheckDecryptor18 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptor18", DefaultCheckDecryptor18))
                    .CheckDecryptor19 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptor19", DefaultCheckDecryptor19))
                    .CheckDecryptorUseforT2 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptorUseforT2", DefaultCheckDecryptorUseforT2))
                    .CheckDecryptorUseforT3 = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckDecryptorUseforT3", defaultCheckDecryptorUseforT3))
                    .CheckIgnoreInventionRE = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckIgnoreInventionRE", DefaultCheckIgnoreInventionRE))
                    .CheckRelicWrecked = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRelicWrecked", DefaultCheckRelicWrecked))
                    .CheckRelicIntact = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRelicIntact", DefaultCheckRelicIntact))
                    .CheckRelicMalfunction = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRelicMalfunction", DefaultCheckRelicMalfunction))
                    .CheckOnlyBuild = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckOnlyBuild", DefaultCheckOnlyBuild))
                    .CheckOnlyInvent = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckOnlyInvent", DefaultCheckOnlyInvent))
                    .CheckOnlyRE = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckOnlyRE", DefaultCheckOnlyRE))
                    .CheckIncludeTaxes = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckIncludeTaxes", DefaultCheckIncludeTaxes))
                    .CheckIncludeBrokersFees = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckIncludeBrokersFees", DefaultIncludeBrokersFees))
                    .CheckIncludeUsage = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckIncludeUsage", DefaultCheckIncludeUsage))
                    .CheckRaceAmarr = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRaceAmarr", DefaultCheckRaceAmarr))
                    .CheckRaceCaldari = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRaceCaldari", DefaultCheckRaceCaldari))
                    .CheckRaceGallente = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRaceGallente", DefaultCheckRaceGallente))
                    .CheckRaceMinmatar = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRaceMinmatar", DefaultCheckRaceMinmatar))
                    .CheckRacePirate = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRacePirate", DefaultCheckRacePirate))
                    .CheckRaceOther = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckRaceOther", DefaultCheckRaceOther))
                    .SortBy = CStr(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeString, "ManufacturingSettings", "SortBy", DefaultSortBy))
                    .PriceCompare = CStr(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeString, "ManufacturingSettings", "PriceCompare", DefaultPriceCompare))
                    .CheckIncludeT2Owned = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckIncludeT2Owned", DefaultCheckIncludeT2Owned))
                    .CheckIncludeT3Owned = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckIncludeT3Owned", DefaultCheckIncludeT3Owned))
                    .IgnoreSVRThreshold = CDbl(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeDouble, "ManufacturingSettings", "IgnoreLowSVRThreshold", DefaultIgnoreSVRThreshold))
                    .CheckSVRIncludeNull = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckSVRIncludeNull", DefaultCheckSVRIncludeNull))
                    .AveragePriceRegion = CStr(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeString, "ManufacturingSettings", "AveragePriceRegion", DefaultSVRRegion))
                    .ProductionLines = CInt(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeInteger, "ManufacturingSettings", "ProductionLines", DefaultCalcProductionLines))
                    .LaboratoryLines = CInt(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeInteger, "ManufacturingSettings", "LaboratoryLines", DefaultCalcLaboratoryLines))
                    .Runs = CInt(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeInteger, "ManufacturingSettings", "Runs", DefaultCalcRuns))
                    .CheckSmall = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckSmall", DefaultCalcSizeChecks))
                    .CheckMedium = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckMedium", DefaultCalcSizeChecks))
                    .CheckLarge = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckLarge", DefaultCalcSizeChecks))
                    .CheckXL = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckXL", DefaultCalcSizeChecks))
                    .CheckCapitalComponentsFacility = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckCapitalComponentsFacility", DefaultCheckT3Destroyers))
                    .CheckT3DestroyerFacility = CBool(GetSettingValue(ManufacturingSettingsFileName, SettingTypes.TypeBoolean, "ManufacturingSettings", "CheckT3DestroyerFacility", DefaultCheckCapComponents))
                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultManufacturingSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Manufacturing Tab Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultManufacturingSettings()
        End Try

        ' Save them locally and then export
        ManufacturingSettings = TempSettings

        Return TempSettings

    End Function

    ' Loads the Defaults for the tab
    Public Function SetDefaultManufacturingSettings() As ManufacturingTabSettings
        Dim LocalSettings As ManufacturingTabSettings

        With LocalSettings
            .BlueprintType = DefaultBlueprintType
            .CheckTech1 = DefaultCheckTech1
            .CheckTech2 = DefaultCheckTech2
            .CheckTech3 = DefaultCheckTech3
            .CheckTechStoryline = DefaultCheckTechStoryline
            .CheckTechNavy = DefaultCheckTechNavy
            .CheckTechPirate = DefaultCheckTechPirate
            .ItemTypeFilter = DefaultItemTypeFilter
            .TextItemFilter = DefaultTextItemFilter
            .CheckBPTypeShips = DefaultCheckBPTypeShips
            .CheckBPTypeDrones = DefaultCheckBPTypeDrones
            .CheckBPTypeComponents = DefaultCheckBPTypeComponents
            .CheckBPTypeStructures = DefaultCheckBPTypeStructures
            .CheckBPTypeMisc = DefaultCheckBPTypeTools
            .CheckBPTypeModules = DefaultCheckBPTypeModules
            .CheckBPTypeAmmoCharges = DefaultCheckBPTypeAmmoCharges
            .CheckBPTypeRigs = DefaultCheckBPTypeRigs
            .CheckBPTypeSubsystems = DefaultCheckBPTypeSubsystems
            .CheckBPTypeBoosters = DefaultCheckBPTypeBoosters
            .CheckBPTypeCelestials = DefaultCheckBPTypeCelestials
            .CheckBPTypeStationParts = DefaultCheckBPTypeStationParts
            .CheckBPTypeDeployables = DefaultCheckBPTypeDeployables
            .AveragePriceDuration = DefaultAveragePriceDuration
            .CheckDecryptorNone = DefaultCheckDecryptorNone
            .CheckDecryptor06 = DefaultCheckDecryptor06
            .CheckDecryptor09 = DefaultCheckDecryptor09
            .CheckDecryptor10 = DefaultCheckDecryptor10
            .CheckDecryptor11 = DefaultCheckDecryptor11
            .CheckDecryptor12 = DefaultCheckDecryptor12
            .CheckDecryptor15 = DefaultCheckDecryptor15
            .CheckDecryptor18 = DefaultCheckDecryptor18
            .CheckDecryptor19 = DefaultCheckDecryptor19
            .CheckDecryptorUseforT2 = DefaultCheckDecryptorUseforT2
            .CheckDecryptorUseforT3 = defaultCheckDecryptorUseforT3
            .CheckIgnoreInventionRE = DefaultCheckIgnoreInventionRE
            .CheckRelicWrecked = DefaultCheckRelicWrecked
            .CheckRelicIntact = DefaultCheckRelicIntact
            .CheckRelicMalfunction = DefaultCheckRelicMalfunction
            .CheckOnlyBuild = DefaultCheckOnlyBuild
            .CheckOnlyInvent = DefaultCheckOnlyInvent
            .CheckOnlyRE = DefaultCheckOnlyRE
            .CheckIncludeTaxes = DefaultCheckIncludeTaxes
            .CheckIncludeBrokersFees = DefaultIncludeBrokersFees
            .CheckIncludeUsage = DefaultCheckIncludeUsage
            .CheckRaceAmarr = DefaultCheckRaceAmarr
            .CheckRaceCaldari = DefaultCheckRaceCaldari
            .CheckRaceGallente = DefaultCheckRaceGallente
            .CheckRaceMinmatar = DefaultCheckRaceMinmatar
            .CheckRacePirate = DefaultCheckRacePirate
            .CheckRaceOther = DefaultCheckRaceOther
            .SortBy = DefaultSortBy
            .PriceCompare = DefaultPriceCompare
            .CheckIncludeT2Owned = DefaultCheckIncludeT2Owned
            .CheckIncludeT3Owned = DefaultCheckIncludeT3Owned
            .IgnoreSVRThreshold = DefaultIgnoreSVRThreshold
            .CheckSVRIncludeNull = DefaultCheckSVRIncludeNull
            .AveragePriceRegion = DefaultSVRRegion
            .ProductionLines = DefaultCalcProductionLines
            .LaboratoryLines = DefaultCalcLaboratoryLines
            .Runs = DefaultCalcRuns
            .CheckSmall = DefaultCalcSizeChecks
            .CheckMedium = DefaultCalcSizeChecks
            .CheckLarge = DefaultCalcSizeChecks
            .CheckXL = DefaultCalcSizeChecks
            .CheckT3DestroyerFacility = DefaultCheckT3Destroyers
            .CheckCapitalComponentsFacility = DefaultCheckCapComponents
        End With

        ' Save locally
        ManufacturingSettings = LocalSettings
        Return LocalSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveManufacturingSettings(SentSettings As ManufacturingTabSettings)
        Dim ManufacturingSettingsList(65) As Setting

        Try
            ManufacturingSettingsList(0) = New Setting("BlueprintType", CStr(SentSettings.BlueprintType))
            ManufacturingSettingsList(1) = New Setting("CheckTech1", CStr(SentSettings.CheckTech1))
            ManufacturingSettingsList(2) = New Setting("CheckTech2", CStr(SentSettings.CheckTech2))
            ManufacturingSettingsList(3) = New Setting("CheckTech3", CStr(SentSettings.CheckTech3))
            ManufacturingSettingsList(4) = New Setting("CheckTechStoryline", CStr(SentSettings.CheckTechStoryline))
            ManufacturingSettingsList(5) = New Setting("CheckTechNavy", CStr(SentSettings.CheckTechNavy))
            ManufacturingSettingsList(6) = New Setting("CheckTechPirate", CStr(SentSettings.CheckTechPirate))
            ManufacturingSettingsList(7) = New Setting("ItemTypeFilter", CStr(SentSettings.ItemTypeFilter))
            ManufacturingSettingsList(8) = New Setting("TextItemFilter", CStr(SentSettings.TextItemFilter))
            ManufacturingSettingsList(9) = New Setting("CheckBPTypeShips", CStr(SentSettings.CheckBPTypeShips))
            ManufacturingSettingsList(10) = New Setting("CheckBPTypeDrones", CStr(SentSettings.CheckBPTypeDrones))
            ManufacturingSettingsList(11) = New Setting("CheckBPTypeComponents", CStr(SentSettings.CheckBPTypeComponents))
            ManufacturingSettingsList(12) = New Setting("CheckBPTypeStructures", CStr(SentSettings.CheckBPTypeStructures))
            ManufacturingSettingsList(13) = New Setting("CheckBPTypeMisc", CStr(SentSettings.CheckBPTypeMisc))
            ManufacturingSettingsList(14) = New Setting("CheckBPTypeModules", CStr(SentSettings.CheckBPTypeModules))
            ManufacturingSettingsList(15) = New Setting("CheckBPTypeAmmoCharges", CStr(SentSettings.CheckBPTypeAmmoCharges))
            ManufacturingSettingsList(16) = New Setting("CheckBPTypeRigs", CStr(SentSettings.CheckBPTypeRigs))
            ManufacturingSettingsList(17) = New Setting("CheckBPTypeSubsystems", CStr(SentSettings.CheckBPTypeSubsystems))
            ManufacturingSettingsList(18) = New Setting("CheckBPTypeBoosters", CStr(SentSettings.CheckBPTypeBoosters))
            ManufacturingSettingsList(19) = New Setting("AveragePriceDuration", CStr(SentSettings.AveragePriceDuration))
            ManufacturingSettingsList(20) = New Setting("CheckDecryptorNone", CStr(SentSettings.CheckDecryptorNone))
            ManufacturingSettingsList(21) = New Setting("CheckDecryptor06", CStr(SentSettings.CheckDecryptor06))
            ManufacturingSettingsList(22) = New Setting("CheckDecryptor10", CStr(SentSettings.CheckDecryptor10))
            ManufacturingSettingsList(23) = New Setting("CheckDecryptor11", CStr(SentSettings.CheckDecryptor11))
            ManufacturingSettingsList(24) = New Setting("CheckDecryptor12", CStr(SentSettings.CheckDecryptor12))
            ManufacturingSettingsList(25) = New Setting("CheckDecryptor18", CStr(SentSettings.CheckDecryptor18))
            ManufacturingSettingsList(26) = New Setting("CheckIgnoreInventionRE", CStr(SentSettings.CheckIgnoreInventionRE))
            ManufacturingSettingsList(27) = New Setting("CheckRelicWrecked", CStr(SentSettings.CheckRelicWrecked))
            ManufacturingSettingsList(28) = New Setting("CheckRelicIntact", CStr(SentSettings.CheckRelicIntact))
            ManufacturingSettingsList(29) = New Setting("CheckRelicMalfunction", CStr(SentSettings.CheckRelicMalfunction))
            ManufacturingSettingsList(30) = New Setting("CheckOnlyBuild", CStr(SentSettings.CheckOnlyBuild))
            ManufacturingSettingsList(31) = New Setting("CheckOnlyInvent", CStr(SentSettings.CheckOnlyInvent))
            ManufacturingSettingsList(32) = New Setting("CheckOnlyRE", CStr(SentSettings.CheckOnlyRE))
            ManufacturingSettingsList(33) = New Setting("CheckIncludeTaxes", CStr(SentSettings.CheckIncludeTaxes))
            ManufacturingSettingsList(34) = New Setting("CheckIncludeUsage", CStr(SentSettings.CheckIncludeUsage))
            ManufacturingSettingsList(35) = New Setting("CheckRaceAmarr", CStr(SentSettings.CheckRaceAmarr))
            ManufacturingSettingsList(36) = New Setting("CheckRaceCaldari", CStr(SentSettings.CheckRaceCaldari))
            ManufacturingSettingsList(37) = New Setting("CheckRaceGallente", CStr(SentSettings.CheckRaceGallente))
            ManufacturingSettingsList(38) = New Setting("CheckRaceMinmatar", CStr(SentSettings.CheckRaceMinmatar))
            ManufacturingSettingsList(39) = New Setting("CheckRacePirate", CStr(SentSettings.CheckRacePirate))
            ManufacturingSettingsList(40) = New Setting("CheckRaceOther", CStr(SentSettings.CheckRaceOther))
            ManufacturingSettingsList(41) = New Setting("SortBy", CStr(SentSettings.SortBy))
            ManufacturingSettingsList(42) = New Setting("PriceCompare", CStr(SentSettings.PriceCompare))
            ManufacturingSettingsList(43) = New Setting("CheckIncludeT2Owned", CStr(SentSettings.CheckIncludeT2Owned))
            ManufacturingSettingsList(44) = New Setting("CheckIncludeT3Owned", CStr(SentSettings.CheckIncludeT3Owned))
            ManufacturingSettingsList(45) = New Setting("IgnoreLowSVRThreshold", CStr(SentSettings.IgnoreSVRThreshold))
            ManufacturingSettingsList(46) = New Setting("CheckSVRIncludeNull", CStr(SentSettings.CheckSVRIncludeNull))
            ManufacturingSettingsList(47) = New Setting("AveragePriceRegion", CStr(SentSettings.AveragePriceRegion))
            ManufacturingSettingsList(48) = New Setting("ProductionLines", CStr(SentSettings.ProductionLines))
            ManufacturingSettingsList(49) = New Setting("LaboratoryLines", CStr(SentSettings.LaboratoryLines))
            ManufacturingSettingsList(50) = New Setting("CheckDecryptor09", CStr(SentSettings.CheckDecryptor09))
            ManufacturingSettingsList(51) = New Setting("CheckDecryptor15", CStr(SentSettings.CheckDecryptor15))
            ManufacturingSettingsList(52) = New Setting("CheckDecryptor19", CStr(SentSettings.CheckDecryptor19))
            ManufacturingSettingsList(53) = New Setting("Runs", CStr(SentSettings.Runs))
            ManufacturingSettingsList(54) = New Setting("CheckBPTypeCelestials", CStr(SentSettings.CheckBPTypeCelestials))
            ManufacturingSettingsList(55) = New Setting("CheckBPTypeDeployables", CStr(SentSettings.CheckBPTypeDeployables))
            ManufacturingSettingsList(56) = New Setting("CheckSmall", CStr(SentSettings.CheckSmall))
            ManufacturingSettingsList(57) = New Setting("CheckMedium", CStr(SentSettings.CheckMedium))
            ManufacturingSettingsList(58) = New Setting("CheckLarge", CStr(SentSettings.CheckLarge))
            ManufacturingSettingsList(59) = New Setting("CheckXL", CStr(SentSettings.CheckXL))
            ManufacturingSettingsList(60) = New Setting("CheckBPTypeStationParts", CStr(SentSettings.CheckBPTypeStationParts))
            ManufacturingSettingsList(61) = New Setting("CheckIncludeBrokersFees", CStr(SentSettings.CheckIncludeBrokersFees))
            ManufacturingSettingsList(62) = New Setting("CheckDecryptorUseforT2", CStr(SentSettings.CheckDecryptorUseforT2))
            ManufacturingSettingsList(63) = New Setting("CheckDecryptorUseforT3", CStr(SentSettings.CheckDecryptorUseforT3))
            ManufacturingSettingsList(64) = New Setting("CheckCapitalComponentsFacility", CStr(SentSettings.CheckCapitalComponentsFacility))
            ManufacturingSettingsList(65) = New Setting("CheckT3DestroyerFacility", CStr(SentSettings.CheckT3DestroyerFacility))

            Call WriteSettingsToFile(ManufacturingSettingsFileName, ManufacturingSettingsList, "ManufacturingSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Manufacturing Tab Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetManufacturingSettings() As ManufacturingTabSettings
        Return ManufacturingSettings
    End Function

#End Region

#Region "Datacore Tab Settings"

    ' Loads the tab settings
    Public Function LoadDatacoreSettings() As DataCoreTabSettings
        Dim TempSettings As DataCoreTabSettings = Nothing

        Try

            ' Dim the settings
            ReDim TempSettings.SkillsLevel(NumberofDCSettingsSkillRecords)
            ReDim TempSettings.SkillsChecked(NumberofDCSettingsSkillRecords)
            ReDim TempSettings.CorpsStanding(NumberofDCSettingsCorpRecords)
            ReDim TempSettings.CorpsChecked(NumberofDCSettingsCorpRecords)

            If File.Exists(SettingsFolder & DatacoreSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .PricesFrom = CStr(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeString, "DatacoreSettings", "PricesFrom", DefaultReactPOSFuelCost))
                    .CheckHighSecAgents = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckHighSecAgents", DefaultReactCheckTaxes))
                    .CheckLowNullSecAgents = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckLowNullSecAgents", DefaultReactCheckFees))
                    .CheckIncludeAgentsCannotAccess = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckIncludeAgentsCannotAccess", DefaultReactItemChecks))
                    .AgentsInRegion = CStr(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeString, "DatacoreSettings", "AgentsInRegion", DefaultReactItemChecks))
                    .CheckSovAmarr = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckSovAmarr", DefaultReactItemChecks))
                    .CheckSovAmmatar = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckSovAmmatar", DefaultReactItemChecks))
                    .CheckSovGallente = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckSovGallente", DefaultReactItemChecks))
                    .CheckSovSyndicate = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckSovSyndicate", DefaultReactItemChecks))
                    .CheckSovKhanid = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckSovKhanid", DefaultReactItemChecks))
                    .CheckSovThukker = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckSovThukker", DefaultReactItemChecks))
                    .CheckSovCaldari = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckSovCaldari", DefaultReactItemChecks))
                    .CheckSovMinmatar = CBool(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeBoolean, "DatacoreSettings", "CheckSovMinmatar", DefaultReactItemChecks))

                    For i = 1 To 17
                        .SkillsChecked(i - 1) = CInt(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeInteger, "DatacoreSettings", "Skill" & CStr(i) & "Checked", DefaultSkillLevelChecked))
                        .SkillsLevel(i - 1) = CInt(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeInteger, "DatacoreSettings", "Skill" & CStr(i) & "Level ", DefaultSkillLevel))
                    Next

                    For i = 1 To 13
                        .CorpsChecked(i - 1) = CInt(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeInteger, "DatacoreSettings", "Corp" & CStr(i) & "Checked", DefaultSkillLevelChecked))
                        .CorpsStanding(i - 1) = CInt(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeInteger, "DatacoreSettings", "Corp" & CStr(i) & "Standing ", DefaultSkillLevel))
                    Next

                    .Negotiation = CInt(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeInteger, "DatacoreSettings", "Negotiation", DefaultNegotiation))
                    .Connections = CInt(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeInteger, "DatacoreSettings", "Connections", DefaultConnections))
                    .ResearchProjectMgt = CInt(GetSettingValue(DatacoreSettingsFileName, SettingTypes.TypeInteger, "DatacoreSettings", "ResearchProjectMgt", DefaultResearchProjMgt))

                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultDatacoreSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Datacore Tab Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultDatacoreSettings()
        End Try

        ' Save them locally and then export
        DatacoreSettings = TempSettings

        Return TempSettings

    End Function

    Public Function SetDefaultDatacoreSettings() As DataCoreTabSettings
        Dim LocalSettings As DataCoreTabSettings

        ReDim LocalSettings.SkillsChecked(NumberofDCSettingsSkillRecords)
        ReDim LocalSettings.SkillsLevel(NumberofDCSettingsSkillRecords)

        ReDim LocalSettings.CorpsChecked(NumberofDCSettingsCorpRecords)
        ReDim LocalSettings.CorpsStanding(NumberofDCSettingsCorpRecords)

        With LocalSettings
            .PricesFrom = DefaultDCPricesFrom
            .CheckHighSecAgents = DefaultDCCheckHighSec
            .CheckLowNullSecAgents = DefaultDCCheckLowNullSec
            .CheckIncludeAgentsCannotAccess = DefaultDCIncludeAgentsCantUse
            .AgentsInRegion = DefaultDCAgentsInRegion
            .CheckSovAmarr = DefaultDCSovCheck
            .CheckSovAmmatar = DefaultDCSovCheck
            .CheckSovGallente = DefaultDCSovCheck
            .CheckSovSyndicate = DefaultDCSovCheck
            .CheckSovKhanid = DefaultDCSovCheck
            .CheckSovThukker = DefaultDCSovCheck
            .CheckSovCaldari = DefaultDCSovCheck
            .CheckSovMinmatar = DefaultDCSovCheck

            For i = 0 To .SkillsChecked.Count - 1
                .SkillsChecked(i) = DefaultSkillLevelChecked
                .SkillsLevel(i) = DefaultSkillLevel
            Next

            For i = 0 To .CorpsChecked.Count - 1
                .CorpsChecked(i) = DefaultCorpStandingChecked
                .CorpsStanding(i) = DefaultCorpStanding
            Next

            .Negotiation = DefaultNegotiation
            .Connections = DefaultConnections
            .ResearchProjectMgt = DefaultResearchProjMgt

        End With
        ' Save locally
        DatacoreSettings = LocalSettings
        Return LocalSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveDatacoreSettings(SentSettings As DataCoreTabSettings)
        Dim DatacoreSettingsList(75) As Setting
        Dim j As Integer

        Try
            DatacoreSettingsList(0) = New Setting("PricesFrom", CStr(SentSettings.PricesFrom))
            DatacoreSettingsList(1) = New Setting("CheckHighSecAgents", CStr(SentSettings.CheckHighSecAgents))
            DatacoreSettingsList(2) = New Setting("CheckLowNullSecAgents", CStr(SentSettings.CheckLowNullSecAgents))
            DatacoreSettingsList(3) = New Setting("CheckIncludeAgentsCannotAccess", CStr(SentSettings.CheckIncludeAgentsCannotAccess))
            DatacoreSettingsList(4) = New Setting("AgentsInRegion", CStr(SentSettings.AgentsInRegion))
            DatacoreSettingsList(5) = New Setting("CheckSovAmarr", CStr(SentSettings.CheckSovAmarr))
            DatacoreSettingsList(6) = New Setting("CheckSovAmmatar", CStr(SentSettings.CheckSovAmmatar))
            DatacoreSettingsList(7) = New Setting("CheckSovGallente", CStr(SentSettings.CheckSovGallente))
            DatacoreSettingsList(8) = New Setting("CheckSovSyndicate", CStr(SentSettings.CheckSovSyndicate))
            DatacoreSettingsList(9) = New Setting("CheckSovKhanid", CStr(SentSettings.CheckSovKhanid))
            DatacoreSettingsList(10) = New Setting("CheckSovThukker", CStr(SentSettings.CheckSovThukker))
            DatacoreSettingsList(11) = New Setting("CheckSovCaldari", CStr(SentSettings.CheckSovCaldari))
            DatacoreSettingsList(12) = New Setting("CheckSovMinmatar", CStr(SentSettings.CheckSovMinmatar))

            ' Skills
            j = 0
            For i = 13 To 29
                j += 1
                DatacoreSettingsList(i) = New Setting("Skill" & CStr(j) & "Level", CStr(SentSettings.SkillsLevel(j - 1)))
            Next

            j = 0
            For i = 30 To 46
                j += 1
                DatacoreSettingsList(i) = New Setting("Skill" & CStr(j) & "Checked", CStr(SentSettings.SkillsChecked(j - 1)))
            Next

            ' Corp Standings
            j = 0
            For i = 47 To 59
                j += 1
                DatacoreSettingsList(i) = New Setting("Corp" & CStr(j) & "Standing", CStr(SentSettings.CorpsStanding(j - 1)))
            Next

            j = 0
            For i = 60 To 72
                j += 1
                DatacoreSettingsList(i) = New Setting("Corp" & CStr(j) & "Checked", CStr(SentSettings.CorpsChecked(j - 1)))
            Next

            DatacoreSettingsList(73) = New Setting("Negotiation", CStr(SentSettings.Negotiation))
            DatacoreSettingsList(74) = New Setting("Connections", CStr(SentSettings.Connections))
            DatacoreSettingsList(75) = New Setting("ResearchProjectMgt", CStr(SentSettings.ResearchProjectMgt))

            Call WriteSettingsToFile(DatacoreSettingsFileName, DatacoreSettingsList, "DatacoreSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Datacore Tab Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetDatacoreSettings() As DataCoreTabSettings
        Return DatacoreSettings
    End Function

#End Region

#Region "Reactions Tab Settings"

    ' Loads the tab settings
    Public Function LoadReactionSettings() As ReactionsTabSettings
        Dim TempSettings As ReactionsTabSettings = Nothing

        Try
            If File.Exists(SettingsFolder & ReactionSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .POSFuelCost = CDbl(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeDouble, "ReactionSettings", "POSFuelCost", DefaultReactPOSFuelCost))
                    .CheckTaxes = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckTaxes", DefaultReactCheckTaxes))
                    .CheckFees = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckFees", DefaultReactCheckFees))
                    .CheckAdvMoonMats = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckAdvMoonMats", DefaultReactItemChecks))
                    .CheckProcessedMoonMats = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckProcessedMoonMats", DefaultReactItemChecks))
                    .CheckHybrid = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckHybrid", DefaultReactItemChecks))
                    .CheckComplexBio = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckComplexBio", DefaultReactItemChecks))
                    .CheckSimpleBio = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckSimpleBio", DefaultReactItemChecks))
                    .CheckBuildBasic = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckBuildBasic", DefaultReactItemChecks))
                    .CheckIgnoreMarket = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckIgnoreMarket", DefaultReactItemChecks))
                    .CheckRefine = CBool(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeBoolean, "ReactionSettings", "CheckRefine", DefaultReactItemChecks))
                    .NumberofPOS = CInt(GetSettingValue(ReactionSettingsFileName, SettingTypes.TypeInteger, "ReactionSettings", "NumberofPOS", DefaultReactNumPOS))
                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultReactionSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Reaction Tab Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultReactionSettings()
        End Try

        ' Save them locally and then export
        ReactionSettings = TempSettings

        Return TempSettings

    End Function

    ' Loads the Defaults for the tab
    Public Function SetDefaultReactionSettings() As ReactionsTabSettings
        Dim LocalSettings As ReactionsTabSettings

        LocalSettings.POSFuelCost = DefaultReactPOSFuelCost
        LocalSettings.CheckTaxes = DefaultReactCheckTaxes
        LocalSettings.CheckFees = DefaultReactCheckFees
        LocalSettings.CheckAdvMoonMats = DefaultReactItemChecks
        LocalSettings.CheckProcessedMoonMats = DefaultReactItemChecks
        LocalSettings.CheckHybrid = DefaultReactItemChecks
        LocalSettings.CheckComplexBio = DefaultReactItemChecks
        LocalSettings.CheckSimpleBio = DefaultReactItemChecks
        LocalSettings.CheckBuildBasic = DefaultReactItemChecks
        LocalSettings.CheckIgnoreMarket = DefaultReactItemChecks
        LocalSettings.CheckRefine = DefaultReactItemChecks
        LocalSettings.NumberofPOS = DefaultReactNumPOS

        ' Save locally
        ReactionSettings = LocalSettings
        Return LocalSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveReactionSettings(SentSettings As ReactionsTabSettings)
        Dim ReactionSettingsList(11) As Setting

        Try
            ReactionSettingsList(0) = New Setting("POSFuelCost", CStr(SentSettings.POSFuelCost))
            ReactionSettingsList(1) = New Setting("CheckTaxes", CStr(SentSettings.CheckTaxes))
            ReactionSettingsList(2) = New Setting("CheckFees", CStr(SentSettings.CheckFees))
            ReactionSettingsList(3) = New Setting("CheckAdvMoonMats", CStr(SentSettings.CheckAdvMoonMats))
            ReactionSettingsList(4) = New Setting("CheckProcessedMoonMats", CStr(SentSettings.CheckProcessedMoonMats))
            ReactionSettingsList(5) = New Setting("CheckHybrid", CStr(SentSettings.CheckHybrid))
            ReactionSettingsList(6) = New Setting("CheckComplexBio", CStr(SentSettings.CheckComplexBio))
            ReactionSettingsList(7) = New Setting("CheckSimpleBio", CStr(SentSettings.CheckSimpleBio))
            ReactionSettingsList(8) = New Setting("CheckBuildBasic", CStr(SentSettings.CheckBuildBasic))
            ReactionSettingsList(9) = New Setting("CheckIgnoreMarket", CStr(SentSettings.CheckIgnoreMarket))
            ReactionSettingsList(10) = New Setting("CheckRefine", CStr(SentSettings.CheckRefine))
            ReactionSettingsList(11) = New Setting("NumberofPOS", CStr(SentSettings.NumberofPOS))

            Call WriteSettingsToFile(ReactionSettingsFileName, ReactionSettingsList, "ReactionSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Reaction Tab Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetReactionSettings() As ReactionsTabSettings
        Return ReactionSettings
    End Function

#End Region

#Region "Mining Tab Settings"

    ' Loads the tab settings
    Public Function LoadMiningSettings() As MiningTabSettings
        Dim TempSettings As MiningTabSettings = Nothing

        Try
            If File.Exists(SettingsFolder & MiningSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .OreType = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "OreType", DefaultMiningOreType))
                    .CheckHighYieldOres = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckHighYieldOres", DefaultMiningCheckHighYieldOres))
                    .CheckHighSecOres = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckHighSecOres", DefaultMiningCheckHighSecOres))
                    .CheckLowSecOres = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckLowSecOres", DefaultMiningCheckLowSecOres))
                    .CheckNullSecOres = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckNullSecOres", DefaultMiningCheckNullSecOres))
                    .CheckSovAmarr = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovAmarr", DefaultMiningCheckSovAmarr))
                    .CheckSovCaldari = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovCaldari", DefaultMiningCheckSovCaldari))
                    .CheckSovGallente = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovGallente", DefaultMiningCheckSovGallente))
                    .CheckSovMinmatar = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovMinmatar", DefaultMiningCheckSovMinmatar))
                    .CheckIncludeFees = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckIncludeFees", DefaultMiningCheckIncludeFees))
                    .CheckIncludeTaxes = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckIncludeTaxes", DefaultMiningCheckIncludeTaxes))
                    .CheckIncludeJumpFuelCosts = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckIncludeJumpFuelCosts", DefaultMiningCheckIncludeJumpFuelCosts))
                    .TotalJumpFuelCost = CDbl(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeDouble, "MiningSettings", "TotalJumpFuelCost", DefaultMiningTotalJumpFuelCost))
                    .TotalJumpFuelM3 = CDbl(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeDouble, "MiningSettings", "TotalJumpFuelM3", DefaultMiningTotalJumpFuelM3))
                    .JumpCompressedOre = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "JumpCompressedOre", DefaultMiningJumpCompressedOre))
                    .JumpMinerals = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "JumpMinerals", DefaultMiningJumpMinerals))
                    .OreMiningShip = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "OreMiningShip", DefaultMiningMiningShip))
                    .IceMiningShip = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "IceMiningShip", DefaultMiningIceMiningShip))
                    .GasMiningShip = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "GasMiningShip", DefaultMiningGasMiningShip))
                    .OreStrip = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "OreStrip", DefaultMiningOreStrip))
                    .IceStrip = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "IceStrip", DefaultMiningIceStrip))
                    .GasHarvester = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "GasHarvester", DefaultMiningGasHarvester))
                    .NumOreMiners = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "NumOreMiners", DefaultMiningNumOreMiners))
                    .NumIceMiners = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "NumIceMiners", DefaultMiningNumIceMiners))
                    .NumGasHarvesters = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "NumGasHarvesters", DefaultMiningNumGasHarvesters))
                    .OreUpgrade = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "OreUpgrade", DefaultMiningOreUpgrade))
                    .IceUpgrade = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "IceUpgrade", DefaultMiningIceUpgrade))
                    .GasUpgrade = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "GasUpgrade", DefaultMiningGasUpgrade))
                    .NumOreUpgrades = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "NumOreUpgrades", DefaultMiningNumOreUpgrades))
                    .NumIceUpgrades = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "NumIceUpgrades", DefaultMiningNumIceUpgrades))
                    .NumGasUpgrades = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "NumGasUpgrades", DefaultMiningNumGasUpgrades))
                    .MichiiImplant = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "MichiiImplant", DefaultMiningMichiiImplant))
                    .T2Crystals = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "T2Crystals", DefaultMiningT2Crystals))
                    .OreImplant = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "OreImplant", DefaultMiningOreImplant))
                    .IceImplant = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "IceImplant", DefaultMiningIceImplant))
                    .GasImplant = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "GasImplant", DefaultMiningGasImplant))
                    .CheckUseHauler = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckUseHauler", DefaultMiningCheckUseHauler))
                    .RoundTripMin = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "RoundTripMin", DefaultMiningRoundTripMin))
                    .RoundTripSec = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "RoundTripSec", DefaultMiningRoundTripSec))
                    .Haulerm3 = CDbl(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "Haulerm3", DefaultMiningHaulerm3))
                    .CheckUseFleetBooster = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckUseFleetBooster", DefaultMiningCheckUseFleetBooster))
                    .BoosterShip = CStr(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeString, "MiningSettings", "BoosterShip", DefaultMiningBoosterShip))
                    .BoosterShipSkill = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "BoosterShipSkill", DefaultMiningBoosterShipSkill))
                    .MiningFormanSkill = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "MiningFormanSkill", DefaultMiningMiningFormanSkill))
                    .MiningDirectorSkill = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "MiningDirectorSkill", DefaultMiningMiningDirectorSkill))
                    .WarfareLinkSpecSkill = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "WarfareLinkSpecSkill", DefaultMiningWarfareLinkSpecSkill))
                    .CheckMineForemanLaserOpBoost = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "CheckMineForemanLaserOpBoost", DefaultMiningCheckMineForemanLaserOpBoost))
                    .CheckMineForemanLaserRangeBoost = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "CheckMineForemanLaserRangeBoost", DefaultMiningCheckMineForemanLaserOpBoost))
                    .CheckMiningForemanMindLink = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckMiningForemanMindLink", DefaultMiningCheckMiningForemanMindLink))
                    .CheckRorqDeployed = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckRorqDeployed", DefaultMiningRorqDeployed))
                    .MiningDroneM3perHour = CDbl(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeDouble, "MiningSettings", "MiningDroneM3perHour", DefaultMiningDroneM3perHour))
                    .RefineOre = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "RefineOre", DefaultMiningRefineOre))
                    .IndustrialReconfig = CInt(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeInteger, "MiningSettings", "IndustrialReconfig", DefaultIndustrialReconfig))
                    .MercoxitMiningRig = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "MercoxitMiningRig", DefaultMiningRig))
                    .IceMiningRig = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "IceMiningRig", DefaultMiningRig))
                    .CheckSovWormhole = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovWormhole", DefaultMiningCheckSovWormhole))
                    .CheckSovC1 = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovC1", DefaultMiningCheckSovC1))
                    .CheckSovC2 = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovC2", DefaultMiningCheckSovC2))
                    .CheckSovC3 = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovC3", DefaultMiningCheckSovC3))
                    .CheckSovC4 = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovC4", DefaultMiningCheckSovC4))
                    .CheckSovC5 = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovC5", DefaultMiningCheckSovC5))
                    .CheckSovC6 = CBool(GetSettingValue(MiningSettingsFileName, SettingTypes.TypeBoolean, "MiningSettings", "CheckSovC6", DefaultMiningCheckSovC6))
                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultMiningSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Reaction Tab Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultMiningSettings()
        End Try

        ' Save them locally and then export
        MiningSettings = TempSettings

        Return TempSettings

    End Function

    ' Loads the Defaults for the tab
    Public Function SetDefaultMiningSettings() As MiningTabSettings
        Dim LocalSettings As MiningTabSettings

        With LocalSettings
            .OreType = DefaultMiningOreType
            .CheckHighYieldOres = DefaultMiningCheckHighYieldOres
            .CheckHighSecOres = DefaultMiningCheckHighSecOres
            .CheckLowSecOres = DefaultMiningCheckLowSecOres
            .CheckNullSecOres = DefaultMiningCheckNullSecOres
            .CheckSovAmarr = DefaultMiningCheckSovAmarr
            .CheckSovCaldari = DefaultMiningCheckSovCaldari
            .CheckSovGallente = DefaultMiningCheckSovGallente
            .CheckSovMinmatar = DefaultMiningCheckSovMinmatar
            .CheckSovWormhole = DefaultMiningCheckSovWormhole
            .CheckSovC1 = DefaultMiningCheckSovC1
            .CheckSovC2 = DefaultMiningCheckSovC2
            .CheckSovC3 = DefaultMiningCheckSovC3
            .CheckSovC4 = DefaultMiningCheckSovC4
            .CheckSovC5 = DefaultMiningCheckSovC5
            .CheckSovC6 = DefaultMiningCheckSovC6
            .CheckIncludeFees = DefaultMiningCheckIncludeFees
            .CheckIncludeTaxes = DefaultMiningCheckIncludeTaxes
            .CheckIncludeJumpFuelCosts = DefaultMiningCheckIncludeJumpFuelCosts
            .TotalJumpFuelCost = DefaultMiningTotalJumpFuelCost
            .TotalJumpFuelM3 = DefaultMiningTotalJumpFuelM3
            .JumpCompressedOre = DefaultMiningJumpCompressedOre
            .JumpMinerals = DefaultMiningJumpMinerals
            .OreMiningShip = DefaultMiningMiningShip
            .IceMiningShip = DefaultMiningIceMiningShip
            .GasMiningShip = DefaultMiningGasMiningShip
            .OreStrip = DefaultMiningOreStrip
            .IceStrip = DefaultMiningIceStrip
            .GasHarvester = DefaultMiningGasHarvester
            .NumOreMiners = DefaultMiningNumOreMiners
            .NumIceMiners = DefaultMiningNumIceMiners
            .NumGasHarvesters = DefaultMiningNumGasHarvesters
            .OreUpgrade = DefaultMiningOreUpgrade
            .IceUpgrade = DefaultMiningIceUpgrade
            .GasUpgrade = DefaultMiningGasUpgrade
            .NumOreUpgrades = DefaultMiningNumOreUpgrades
            .NumIceUpgrades = DefaultMiningNumIceUpgrades
            .NumGasUpgrades = DefaultMiningNumGasUpgrades
            .MichiiImplant = DefaultMiningMichiiImplant
            .T2Crystals = DefaultMiningT2Crystals
            .OreImplant = DefaultMiningOreImplant
            .IceImplant = DefaultMiningIceImplant
            .GasImplant = DefaultMiningGasImplant
            .CheckUseHauler = DefaultMiningCheckUseHauler
            .RoundTripMin = DefaultMiningRoundTripMin
            .RoundTripSec = DefaultMiningRoundTripSec
            .Haulerm3 = DefaultMiningHaulerm3
            .CheckUseFleetBooster = DefaultMiningCheckUseFleetBooster
            .BoosterShip = DefaultMiningBoosterShip
            .BoosterShipSkill = DefaultMiningBoosterShipSkill
            .MiningFormanSkill = DefaultMiningMiningFormanSkill
            .MiningDirectorSkill = DefaultMiningMiningDirectorSkill
            .WarfareLinkSpecSkill = DefaultMiningWarfareLinkSpecSkill
            .CheckMineForemanLaserOpBoost = DefaultMiningCheckMineForemanLaserOpBoost
            .CheckMineForemanLaserRangeBoost = DefaultMiningCheckMineForemanLaserOpBoost
            .CheckMiningForemanMindLink = DefaultMiningCheckMiningForemanMindLink
            .CheckRorqDeployed = DefaultMiningRorqDeployed
            .MiningDroneM3perHour = DefaultMiningDroneM3perHour
            .RefineOre = DefaultMiningRefineOre
            .IndustrialReconfig = DefaultIndustrialReconfig
            .MercoxitMiningRig = DefaultMiningRig
            .IceMiningRig = DefaultMiningRig
        End With

        ' Save locally
        MiningSettings = LocalSettings
        Return LocalSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveMiningSettings(SentSettings As MiningTabSettings)
        Dim MiningSettingsList(61) As Setting

        Try
            MiningSettingsList(0) = New Setting("OreType", CStr(SentSettings.OreType))
            MiningSettingsList(1) = New Setting("CheckHighYieldOres", CStr(SentSettings.CheckHighYieldOres))
            MiningSettingsList(2) = New Setting("CheckHighSecOres", CStr(SentSettings.CheckHighSecOres))
            MiningSettingsList(3) = New Setting("CheckLowSecOres", CStr(SentSettings.CheckLowSecOres))
            MiningSettingsList(4) = New Setting("CheckNullSecOres", CStr(SentSettings.CheckNullSecOres))
            MiningSettingsList(5) = New Setting("CheckSovAmarr", CStr(SentSettings.CheckSovAmarr))
            MiningSettingsList(6) = New Setting("CheckSovCaldari", CStr(SentSettings.CheckSovCaldari))
            MiningSettingsList(7) = New Setting("CheckSovGallente", CStr(SentSettings.CheckSovGallente))
            MiningSettingsList(8) = New Setting("CheckSovMinmatar", CStr(SentSettings.CheckSovMinmatar))
            MiningSettingsList(9) = New Setting("CheckIncludeFees", CStr(SentSettings.CheckIncludeFees))
            MiningSettingsList(10) = New Setting("CheckIncludeTaxes", CStr(SentSettings.CheckIncludeTaxes))
            MiningSettingsList(11) = New Setting("CheckIncludeJumpFuelCosts", CStr(SentSettings.CheckIncludeJumpFuelCosts))
            MiningSettingsList(12) = New Setting("TotalJumpFuelCost", CStr(SentSettings.TotalJumpFuelCost))
            MiningSettingsList(13) = New Setting("TotalJumpFuelM3", CStr(SentSettings.TotalJumpFuelM3))
            MiningSettingsList(14) = New Setting("JumpCompressedOre", CStr(SentSettings.JumpCompressedOre))
            MiningSettingsList(15) = New Setting("JumpMinerals", CStr(SentSettings.JumpMinerals))
            MiningSettingsList(16) = New Setting("OreMiningShip", CStr(SentSettings.OreMiningShip))
            MiningSettingsList(17) = New Setting("IceMiningShip", CStr(SentSettings.IceMiningShip))
            MiningSettingsList(18) = New Setting("OreStrip", CStr(SentSettings.OreStrip))
            MiningSettingsList(19) = New Setting("IceStrip", CStr(SentSettings.IceStrip))
            MiningSettingsList(20) = New Setting("NumOreMiners", CStr(SentSettings.NumOreMiners))
            MiningSettingsList(21) = New Setting("NumIceMiners", CStr(SentSettings.NumIceMiners))
            MiningSettingsList(22) = New Setting("OreUpgrade", CStr(SentSettings.OreUpgrade))
            MiningSettingsList(23) = New Setting("IceUpgrade", CStr(SentSettings.IceUpgrade))
            MiningSettingsList(24) = New Setting("NumOreUpgrades", CStr(SentSettings.NumOreUpgrades))
            MiningSettingsList(25) = New Setting("NumIceUpgrades", CStr(SentSettings.NumIceUpgrades))
            MiningSettingsList(26) = New Setting("MichiiImplant", CStr(SentSettings.MichiiImplant))
            MiningSettingsList(27) = New Setting("T2Crystals", CStr(SentSettings.T2Crystals))
            MiningSettingsList(28) = New Setting("OreImplant", CStr(SentSettings.OreImplant))
            MiningSettingsList(29) = New Setting("IceImplant", CStr(SentSettings.IceImplant))
            MiningSettingsList(30) = New Setting("CheckUseHauler", CStr(SentSettings.CheckUseHauler))
            MiningSettingsList(31) = New Setting("RoundTripMin", CStr(SentSettings.RoundTripMin))
            MiningSettingsList(32) = New Setting("RoundTripSec", CStr(SentSettings.RoundTripSec))
            MiningSettingsList(33) = New Setting("Haulerm3", CStr(SentSettings.Haulerm3))
            MiningSettingsList(34) = New Setting("CheckUseFleetBooster", CStr(SentSettings.CheckUseFleetBooster))
            MiningSettingsList(35) = New Setting("BoosterShip", CStr(SentSettings.BoosterShip))
            MiningSettingsList(36) = New Setting("BoosterShipSkill", CStr(SentSettings.BoosterShipSkill))
            MiningSettingsList(37) = New Setting("MiningFormanSkill", CStr(SentSettings.MiningFormanSkill))
            MiningSettingsList(38) = New Setting("MiningDirectorSkill", CStr(SentSettings.MiningDirectorSkill))
            MiningSettingsList(39) = New Setting("WarfareLinkSpecSkill", CStr(SentSettings.WarfareLinkSpecSkill))
            MiningSettingsList(40) = New Setting("CheckMineForemanLaserOpBoost", CStr(SentSettings.CheckMineForemanLaserOpBoost))
            MiningSettingsList(41) = New Setting("CheckMiningForemanMindLink", CStr(SentSettings.CheckMiningForemanMindLink))
            MiningSettingsList(42) = New Setting("CheckRorqDeployed", CStr(SentSettings.CheckRorqDeployed))
            MiningSettingsList(43) = New Setting("MiningDroneM3perHour", CStr(SentSettings.MiningDroneM3perHour))
            MiningSettingsList(44) = New Setting("RefineOre", CStr(SentSettings.RefineOre))
            MiningSettingsList(45) = New Setting("IndustrialReconfig", CStr(SentSettings.IndustrialReconfig))
            MiningSettingsList(46) = New Setting("MercoxitMiningRig", CStr(SentSettings.MercoxitMiningRig))
            MiningSettingsList(47) = New Setting("IceMiningRig", CStr(SentSettings.IceMiningRig))
            MiningSettingsList(48) = New Setting("CheckMineForemanLaserRangeBoost", CStr(SentSettings.CheckMineForemanLaserRangeBoost))
            MiningSettingsList(49) = New Setting("GasMiningShip", CStr(SentSettings.GasMiningShip))
            MiningSettingsList(50) = New Setting("GasHarvester", CStr(SentSettings.GasHarvester))
            MiningSettingsList(51) = New Setting("NumGasHarvesters", CStr(SentSettings.NumGasHarvesters))
            MiningSettingsList(52) = New Setting("GasUpgrade", CStr(SentSettings.GasUpgrade))
            MiningSettingsList(53) = New Setting("NumGasUpgrades", CStr(SentSettings.NumGasUpgrades))
            MiningSettingsList(54) = New Setting("GasImplant", CStr(SentSettings.GasImplant))
            MiningSettingsList(55) = New Setting("CheckSovWormhole", CStr(SentSettings.CheckSovWormhole))
            MiningSettingsList(56) = New Setting("CheckSovC1", CStr(SentSettings.CheckSovC1))
            MiningSettingsList(57) = New Setting("CheckSovC2", CStr(SentSettings.CheckSovC2))
            MiningSettingsList(58) = New Setting("CheckSovC3", CStr(SentSettings.CheckSovC3))
            MiningSettingsList(59) = New Setting("CheckSovC4", CStr(SentSettings.CheckSovC4))
            MiningSettingsList(60) = New Setting("CheckSovC5", CStr(SentSettings.CheckSovC5))
            MiningSettingsList(61) = New Setting("CheckSovC6", CStr(SentSettings.CheckSovC6))

            Call WriteSettingsToFile(MiningSettingsFileName, MiningSettingsList, "MiningSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Mining Tab Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetMiningSettings() As MiningTabSettings
        Return MiningSettings
    End Function

#End Region

#Region "Industry Jobs Column Settings"

    ' Loads the tab settings
    Public Function LoadIndustryJobsColumnSettings() As IndustryJobsColumnSettings
        Dim TempSettings As IndustryJobsColumnSettings = Nothing

        Try
            If File.Exists(SettingsFolder & IndustryJobsColumnSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .JobState = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "JobState", DefaultJobState))
                    .TimeToComplete = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "TimeToComplete", DefaultTimeToComplete))
                    .Activity = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "Activity", DefaultActivity))
                    .Status = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "Status", DefaultStatus))
                    .StartTime = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "StartTime", DefaultStartTime))
                    .EndTime = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "EndTime", DefaultEndTime))
                    .CompletionTime = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "CompletionTime", DefaultCompletionTime))
                    .Blueprint = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "Blueprint", DefaultBlueprint))
                    .OutputItem = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "OutputItem", DefaultOutputItem))
                    .OutputItemType = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "OutputItemType", DefaultOutputItemType))
                    .InstallSystem = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "InstallSystem", DefaultInstallSolarSystem))
                    .InstallRegion = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "InstallRegion", DefaultInstallRegion))
                    .LicensedRuns = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "LicensedRuns", DefaultLicensedRuns))
                    .Runs = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "Runs", DefaultRuns))
                    .SuccessfulRuns = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "SuccessfulRuns", DefaultSuccessfulRuns))
                    .BlueprintLocation = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "BlueprintLocation", DefaultBlueprintLocation))
                    .OutputLocation = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "OutputLocation", DefaultOutputLocation))

                    .JobStateWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "JobStateWidth", DefaultIndustryColumnWidth))
                    .TimeToCompleteWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "TimeToCompleteWidth", DefaultIndustryColumnWidth))
                    .ActivityWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "ActivityWidth", DefaultIndustryColumnWidth))
                    .StatusWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "StatusWidth", DefaultIndustryColumnWidth))
                    .StartTimeWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "StartTimewidth", DefaultIndustryColumnWidth))
                    .EndTimeWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "EndTimeWidth", DefaultIndustryColumnWidth))
                    .CompletionTimeWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "CompletionTimeWidth", DefaultIndustryColumnWidth))
                    .BlueprintWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "BlueprintWidth", DefaultIndustryColumnWidth))
                    .OutputItemWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "OutputItemWidth", DefaultIndustryColumnWidth))
                    .OutputItemTypeWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "OutputItemTypeWidth", DefaultIndustryColumnWidth))
                    .InstallSystemWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "InstallSystemWidth", DefaultIndustryColumnWidth))
                    .InstallRegionWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "InstallRegionWidth", DefaultIndustryColumnWidth))
                    .LicensedRunsWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "LiscencedRunsWidth", DefaultIndustryColumnWidth))
                    .RunsWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "RunsWidth", DefaultIndustryColumnWidth))
                    .SuccessfulRunsWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "SuccessfulRunsWidth", DefaultIndustryColumnWidth))
                    .BlueprintLocationWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "BlueprintLocationWidth", DefaultIndustryColumnWidth))
                    .OutputLocationWidth = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "OutputLocationWidth", DefaultIndustryColumnWidth))

                    .OrderByColumn = CInt(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeInteger, "IndustryJobsColumnSettings", "OrderByColumn", DefaultOrderByColumn))
                    .ViewJobType = CStr(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeString, "IndustryJobsColumnSettings", "ViewJobType", DefaultViewJobType))
                    .OrderType = CStr(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeString, "IndustryJobsColumnSettings", "OrderType", DefaultOrderType))
                    .JobTimes = CStr(GetSettingValue(IndustryJobsColumnSettingsFileName, SettingTypes.TypeString, "IndustryJobsColumnSettings", "JobTimes", DefaultJobTimes))

                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultIndustryJobsColumnSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Industry Jobs Column Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultIndustryJobsColumnSettings()
        End Try

        ' Save them locally and then export
        IndustryJobsColumnSettings = TempSettings

        Return TempSettings

    End Function

    ' Loads the Defaults for the tab
    Public Function SetDefaultIndustryJobsColumnSettings() As IndustryJobsColumnSettings
        Dim LocalSettings As IndustryJobsColumnSettings

        With LocalSettings
            .JobState = DefaultJobState
            .TimeToComplete = DefaultTimeToComplete
            .Activity = DefaultActivity
            .Status = DefaultStatus
            .StartTime = DefaultStartTime
            .EndTime = DefaultEndTime
            .CompletionTime = DefaultCompletionTime
            .Blueprint = DefaultBlueprint
            .OutputItem = DefaultOutputItem
            .OutputItemType = DefaultOutputItemType
            .InstallSystem = DefaultInstallSolarSystem
            .InstallRegion = DefaultInstallRegion
            .LicensedRuns = DefaultLicensedRuns
            .Runs = DefaultRuns
            .BlueprintLocation = DefaultBlueprintLocation
            .SuccessfulRuns = DefaultSuccessfulRuns
            .OutputLocation = DefaultOutputLocation

            .JobStateWidth = DefaultIndustryColumnWidth
            .TimeToCompleteWidth = DefaultIndustryColumnWidth
            .ActivityWidth = DefaultIndustryColumnWidth
            .StatusWidth = DefaultIndustryColumnWidth
            .StartTimeWidth = DefaultIndustryColumnWidth
            .EndTimeWidth = DefaultIndustryColumnWidth
            .CompletionTimeWidth = DefaultIndustryColumnWidth
            .BlueprintWidth = DefaultIndustryColumnWidth
            .OutputItemWidth = DefaultIndustryColumnWidth
            .OutputItemTypeWidth = DefaultIndustryColumnWidth
            .InstallSystemWidth = DefaultIndustryColumnWidth
            .InstallRegionWidth = DefaultIndustryColumnWidth
            .LicensedRunsWidth = DefaultIndustryColumnWidth
            .RunsWidth = DefaultIndustryColumnWidth
            .SuccessfulRunsWidth = DefaultIndustryColumnWidth
            .BlueprintLocationWidth = DefaultIndustryColumnWidth
            .OutputLocationWidth = DefaultIndustryColumnWidth

            .OrderByColumn = DefaultOrderByColumn
            .OrderType = DefaultOrderType
            .ViewJobType = DefaultViewJobType
            .JobTimes = DefaultJobTimes
        End With


        ' Save locally
        IndustryJobsColumnSettings = LocalSettings
        Return LocalSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveIndustryJobsColumnSettings(SentSettings As IndustryJobsColumnSettings)
        Dim IndustryJobsColumnSettingsList(37) As Setting

        Try
            IndustryJobsColumnSettingsList(0) = New Setting("JobState", CStr(SentSettings.JobState))
            IndustryJobsColumnSettingsList(1) = New Setting("TimeToComplete", CStr(SentSettings.TimeToComplete))
            IndustryJobsColumnSettingsList(2) = New Setting("Activity", CStr(SentSettings.Activity))
            IndustryJobsColumnSettingsList(3) = New Setting("Status", CStr(SentSettings.Status))
            IndustryJobsColumnSettingsList(4) = New Setting("StartTime", CStr(SentSettings.StartTime))
            IndustryJobsColumnSettingsList(5) = New Setting("EndTime", CStr(SentSettings.EndTime))
            IndustryJobsColumnSettingsList(6) = New Setting("CompletionTime", CStr(SentSettings.CompletionTime))
            IndustryJobsColumnSettingsList(7) = New Setting("Blueprint", CStr(SentSettings.Blueprint))
            IndustryJobsColumnSettingsList(8) = New Setting("OutputItem", CStr(SentSettings.OutputItem))
            IndustryJobsColumnSettingsList(9) = New Setting("OutputItemType", CStr(SentSettings.OutputItemType))
            IndustryJobsColumnSettingsList(10) = New Setting("InstallSystem", CStr(SentSettings.InstallSystem))
            IndustryJobsColumnSettingsList(11) = New Setting("InstallRegion", CStr(SentSettings.InstallRegion))
            IndustryJobsColumnSettingsList(12) = New Setting("LicensedRuns", CStr(SentSettings.LicensedRuns))
            IndustryJobsColumnSettingsList(13) = New Setting("Runs", CStr(SentSettings.Runs))
            IndustryJobsColumnSettingsList(14) = New Setting("SuccessfulRuns", CStr(SentSettings.SuccessfulRuns))
            IndustryJobsColumnSettingsList(15) = New Setting("BlueprintLocation", CStr(SentSettings.BlueprintLocation))
            IndustryJobsColumnSettingsList(16) = New Setting("OutputLocation", CStr(SentSettings.OutputLocation))

            IndustryJobsColumnSettingsList(17) = New Setting("JobStateWidth", CStr(SentSettings.JobStateWidth))
            IndustryJobsColumnSettingsList(18) = New Setting("TimeToCompleteWidth", CStr(SentSettings.TimeToCompleteWidth))
            IndustryJobsColumnSettingsList(19) = New Setting("ActivityWidth", CStr(SentSettings.ActivityWidth))
            IndustryJobsColumnSettingsList(20) = New Setting("StatusWidth", CStr(SentSettings.StatusWidth))
            IndustryJobsColumnSettingsList(21) = New Setting("StartTimeWidth", CStr(SentSettings.StartTimeWidth))
            IndustryJobsColumnSettingsList(22) = New Setting("EndTimeWidth", CStr(SentSettings.EndTimeWidth))
            IndustryJobsColumnSettingsList(23) = New Setting("CompletionTimeWidth", CStr(SentSettings.CompletionTimeWidth))
            IndustryJobsColumnSettingsList(24) = New Setting("BlueprintWidth", CStr(SentSettings.BlueprintWidth))
            IndustryJobsColumnSettingsList(25) = New Setting("OutputItemWidth", CStr(SentSettings.OutputItemWidth))
            IndustryJobsColumnSettingsList(26) = New Setting("OutputItemTypeWidth", CStr(SentSettings.OutputItemTypeWidth))
            IndustryJobsColumnSettingsList(27) = New Setting("InstallSystemWidth", CStr(SentSettings.InstallSystemWidth))
            IndustryJobsColumnSettingsList(28) = New Setting("InstallRegionWidth", CStr(SentSettings.InstallRegionWidth))
            IndustryJobsColumnSettingsList(29) = New Setting("LicensedRunsWidth", CStr(SentSettings.LicensedRunsWidth))
            IndustryJobsColumnSettingsList(30) = New Setting("RunsWidth", CStr(SentSettings.RunsWidth))
            IndustryJobsColumnSettingsList(31) = New Setting("SuccessfulRunsWidth", CStr(SentSettings.SuccessfulRunsWidth))
            IndustryJobsColumnSettingsList(32) = New Setting("BlueprintLocationWidth", CStr(SentSettings.BlueprintLocationWidth))
            IndustryJobsColumnSettingsList(33) = New Setting("OutputLocationWidth", CStr(SentSettings.OutputLocationWidth))

            IndustryJobsColumnSettingsList(34) = New Setting("OrderByColumn", CStr(SentSettings.OrderByColumn))
            IndustryJobsColumnSettingsList(35) = New Setting("ViewJobType", CStr(SentSettings.ViewJobType))
            IndustryJobsColumnSettingsList(36) = New Setting("OrderType", CStr(SentSettings.OrderType))
            IndustryJobsColumnSettingsList(37) = New Setting("JobTimes", CStr(SentSettings.JobTimes))

            Call WriteSettingsToFile(IndustryJobsColumnSettingsFileName, IndustryJobsColumnSettingsList, "IndustryJobsColumnSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Industry Jobs Column Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetIndustryJobsColumnSettings() As IndustryJobsColumnSettings
        Return IndustryJobsColumnSettings
    End Function

#End Region

#Region "Manufacturing Tab Column Settings"

    ' Loads the tab settings
    Public Function LoadManufacturingTabColumnSettings() As ManufacturingTabColumnSettings
        Dim TempSettings As ManufacturingTabColumnSettings = Nothing

        Try
            If File.Exists(SettingsFolder & ManufacturingTabColumnSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .ItemCategory = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ItemCategory", DefaultMTItemCategory))
                    .ItemGroup = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ItemGroup", DefaultMTItemGroup))
                    .ItemName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ItemName", DefaultMTItemName))
                    .Owned = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "Owned", DefaultMTOwned))
                    .Tech = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "Tech", DefaultMTTech))
                    .BPME = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BPME", DefaultMTBPME))
                    .BPTE = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BPTE", DefaultMTBPTE))
                    .Inputs = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "Inputs", DefaultMTInputs))
                    .Compared = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "Compared", DefaultMTCompared))
                    .Runs = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "Runs", DefaultMTRuns))
                    .ProductionLines = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ProductionLines", DefaultMTProductionLines))
                    .LaboratoryLines = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "LaboratoryLines", DefaultMTLaboratoryLines))
                    .TotalInventionRECost = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalInventionRECost", DefaultMTTotalInventionRECost))
                    .TotalCopyCost = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalCopyCost", DefaultMTTotalCopyCost))
                    .TotalManufacturingCost = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalManufacturingCost", DefaultMTTotalManufacturingCost))
                    .Taxes = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "Taxes", DefaultMTTaxes))
                    .BrokerFees = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BrokerFees", DefaultMTBrokerFees))
                    .BPProductionTime = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BPProductionTime", DefaultMTBPProductionTime))
                    .TotalProductionTime = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalProductionTime", DefaultMTTotalProductionTime))
                    .ItemMarketPrice = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ItemMarketPrice", DefaultMTItemMarketPrice))
                    .Profit = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "Profit", DefaultMTProfit))
                    .ProfitPercentage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ProfitPercentage", DefaultMTProfitPercentage))
                    .IskperHour = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "IskperHour", DefaultMTIskperHour))
                    .SVR = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "SVR", DefaultMTSVR))
                    .TotalCost = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalCost", DefaultMTTotalCost))
                    .BaseJobCost = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BaseJobCost", DefaultMTBaseJobCost))
                    .ManufacturingJobFee = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingJobFee", DefaultMTManufacturingJobFee))
                    .ManufacturingFacilityName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityName", DefaultMTManufacturingFacilityName))
                    .ManufacturingFacilitySystem = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilitySystem", DefaultMTManufacturingFacilitySystem))
                    .ManufacturingFacilitySystemIndex = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilitySystemIndex", DefaultMTManufacturingFacilitySystemIndex))
                    .ManufacturingFacilityTax = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityTax", DefaultMTManufacturingFacilityTax))
                    .ManufacturingFacilityRegion = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityRegion", DefaultMTManufacturingFacilityRegion))
                    .ManufacturingFacilityMEBonus = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityMEBonus", DefaultMTManufacturingFacilityMEBonus))
                    .ManufacturingFacilityTEBonus = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityTEBonus", DefaultMTManufacturingFacilityTEBonus))
                    .ManufacturingFacilityUsage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityUsage", DefaultMTManufacturingFacilityUsage))
                    .ComponentFacilityName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityName", DefaultMTComponentFacilityName))
                    .ComponentFacilitySystem = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilitySystem", DefaultMTComponentFacilitySystem))
                    .ComponentFacilitySystemIndex = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilitySystemIndex", DefaultMTComponentFacilitySystemIndex))
                    .ComponentFacilityTax = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityTax", DefaultMTComponentFacilityTax))
                    .ComponentFacilityRegion = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityRegion", DefaultMTComponentFacilityRegion))
                    .ComponentFacilityMEBonus = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityMEBonus", DefaultMTComponentFacilityMEBonus))
                    .ComponentFacilityTEBonus = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityTEBonus", DefaultMTComponentFacilityTEBonus))
                    .ComponentFacilityUsage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityUsage", DefaultMTComponentFacilityUsage))
                    .CopyingFacilityName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityName", DefaultMTCopyingFacilityName))
                    .CopyingFacilitySystem = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilitySystem", DefaultMTCopyingFacilitySystem))
                    .CopyingFacilitySystemIndex = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilitySystemIndex", DefaultMTCopyingFacilitySystemIndex))
                    .CopyingFacilityTax = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityTax", DefaultMTCopyingFacilityTax))
                    .CopyingFacilityRegion = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityRegion", DefaultMTCopyingFacilityRegion))
                    .CopyingFacilityMEBonus = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityMEBonus", DefaultMTCopyingFacilityMEBonus))
                    .CopyingFacilityTEBonus = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityTEBonus", DefaultMTCopyingFacilityTEBonus))
                    .CopyingFacilityUsage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityUsage", DefaultMTCopyingFacilityUsage))
                    .InventionREFacilityName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityName", DefaultMTInventionREFacilityName))
                    .InventionREFacilitySystem = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilitySystem", DefaultMTInventionREFacilitySystem))
                    .InventionREFacilitySystemIndex = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilitySystemIndex", DefaultMTInventionREFacilitySystemIndex))
                    .InventionREFacilityTax = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityTax", DefaultMTInventionREFacilityTax))
                    .InventionREFacilityRegion = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityRegion", DefaultMTInventionREFacilityRegion))
                    .InventionREFacilityMEBonus = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityMEBonus", DefaultMTInventionREFacilityMEBonus))
                    .InventionREFacilityTEBonus = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityTEBonus", DefaultMTInventionREFacilityTEBonus))
                    .InventionREFacilityUsage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityUsage", DefaultMTInventionREFacilityUsage))
                    .ManufacturingTeamName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingTeamName", DefaultMTManufacturingTeamName))
                    .ManufacturingTeamBonuses = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingTeamBonuses", DefaultMTManufacturingTeamBonuses))
                    .ManufacturingTeamUsage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingTeamUsage", DefaultMTManufacturingTeamUsage))
                    .ManufacturingTeamCostModifier = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingTeamCostModifier", DefaultMTManufacturingTeamCostModifier))
                    .ComponentTeamName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentTeamName", DefaultMTComponentTeamName))
                    .ComponentTeamBonuses = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentTeamBonuses", DefaultMTComponentTeamBonuses))
                    .ComponentTeamUsage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentTeamUsage", DefaultMTComponentTeamUsage))
                    .ComponentTeamCostModifier = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentTeamCostModifier", DefaultMTComponentTeamCostModifier))
                    .CopyingTeamName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingTeamName", DefaultMTCopyingTeamName))
                    .CopyingTeamBonuses = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingTeamBonuses", DefaultMTCopyingTeamBonuses))
                    .CopyingTeamUsage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingTeamUsage", DefaultMTCopyingTeamUsage))
                    .CopyingTeamCostModifier = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingTeamCostModifier", DefaultMTCopyingTeamCostModifier))
                    .InventionRETeamName = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionRETeamName", DefaultMTInventionRETeamName))
                    .InventionRETeamBonuses = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionRETeamBonuses", DefaultMTInventionRETeamBonuses))
                    .InventionRETeamUsage = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionRETeamUsage", DefaultMTInventionRETeamUsage))
                    .InventionRETeamCostModifier = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionRETeamCostModifier", DefaultMTInventionRETeamCostModifier))

                    .ItemCategoryWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ItemCategoryWidth", DefaultMTColumnWidth))
                    .ItemGroupWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ItemGroupWidth", DefaultMTColumnWidth))
                    .ItemNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ItemNameWidth", DefaultMTColumnWidth))
                    .OwnedWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "OwnedWidth", DefaultMTColumnWidth))
                    .TechWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TechWidth", DefaultMTColumnWidth))
                    .BPMEWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BPMEWidth", DefaultMTColumnWidth))
                    .BPTEWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BPTEWidth", DefaultMTColumnWidth))
                    .InputsWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InputsWidth", DefaultMTColumnWidth))
                    .ComparedWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComparedWidth", DefaultMTColumnWidth))
                    .RunsWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "RunsWidth", DefaultMTColumnWidth))
                    .ProductionLinesWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ProductionLinesWidth", DefaultMTColumnWidth))
                    .LaboratoryLinesWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "LaboratoryLinesWidth", DefaultMTColumnWidth))
                    .TotalInventionRECostWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalInventionRECostWidth", DefaultMTColumnWidth))
                    .TotalCopyCostWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalCopyCostWidth", DefaultMTColumnWidth))
                    .TotalManufacturingCostWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalManufacturingCostWidth", DefaultMTColumnWidth))
                    .TaxesWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TaxesWidth", DefaultMTColumnWidth))
                    .BrokerFeesWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BrokerFeesWidth", DefaultMTColumnWidth))
                    .BPProductionTimeWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BPProductionTimeWidth", DefaultMTColumnWidth))
                    .TotalProductionTimeWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalProductionTimeWidth", DefaultMTColumnWidth))
                    .ItemMarketPriceWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ItemMarketPriceWidth", DefaultMTColumnWidth))
                    .ProfitWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ProfitWidth", DefaultMTColumnWidth))
                    .ProfitPercentageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ProfitPercentageWidth", DefaultMTColumnWidth))
                    .IskperHourWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "IskperHourWidth", DefaultMTColumnWidth))
                    .SVRWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "SVRWidth", DefaultMTColumnWidth))
                    .TotalCostWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "TotalCostWidth", DefaultMTColumnWidth))
                    .BaseJobCostWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "BaseJobCostWidth", DefaultMTColumnWidth))
                    .ManufacturingJobFeeWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingJobFeeWidth", DefaultMTColumnWidth))
                    .ManufacturingFacilityNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityNameWidth", DefaultMTColumnWidth))
                    .ManufacturingFacilitySystemWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilitySystemWidth", DefaultMTColumnWidth))
                    .ManufacturingFacilitySystemIndexWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilitySystemIndexWidth", DefaultMTColumnWidth))
                    .ManufacturingFacilityTaxWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityTaxWidth", DefaultMTColumnWidth))
                    .ManufacturingFacilityRegionWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityRegionWidth", DefaultMTColumnWidth))
                    .ManufacturingFacilityMEBonusWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityMEBonusWidth", DefaultMTColumnWidth))
                    .ManufacturingFacilityTEBonusWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityTEBonusWidth", DefaultMTColumnWidth))
                    .ManufacturingFacilityUsageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingFacilityUsageWidth", DefaultMTColumnWidth))
                    .ComponentFacilityNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityNameWidth", DefaultMTColumnWidth))
                    .ComponentFacilitySystemWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilitySystemWidth", DefaultMTColumnWidth))
                    .ComponentFacilitySystemIndexWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilitySystemIndexWidth", DefaultMTColumnWidth))
                    .ComponentFacilityTaxWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityTaxWidth", DefaultMTColumnWidth))
                    .ComponentFacilityRegionWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityRegionWidth", DefaultMTColumnWidth))
                    .ComponentFacilityMEBonusWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityMEBonusWidth", DefaultMTColumnWidth))
                    .ComponentFacilityTEBonusWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityTEBonusWidth", DefaultMTColumnWidth))
                    .ComponentFacilityUsageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentFacilityUsageWidth", DefaultMTColumnWidth))
                    .CopyingFacilityNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityNameWidth", DefaultMTColumnWidth))
                    .CopyingFacilitySystemWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilitySystemWidth", DefaultMTColumnWidth))
                    .CopyingFacilitySystemIndexWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilitySystemIndexWidth", DefaultMTColumnWidth))
                    .CopyingFacilityTaxWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityTaxWidth", DefaultMTColumnWidth))
                    .CopyingFacilityRegionWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityRegionWidth", DefaultMTColumnWidth))
                    .CopyingFacilityMEBonusWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityMEBonusWidth", DefaultMTColumnWidth))
                    .CopyingFacilityTEBonusWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityTEBonusWidth", DefaultMTColumnWidth))
                    .CopyingFacilityUsageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingFacilityUsageWidth", DefaultMTColumnWidth))
                    .InventionREFacilityNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityNameWidth", DefaultMTColumnWidth))
                    .InventionREFacilitySystemWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilitySystemWidth", DefaultMTColumnWidth))
                    .InventionREFacilitySystemIndexWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilitySystemIndexWidth", DefaultMTColumnWidth))
                    .InventionREFacilityTaxWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityTaxWidth", DefaultMTColumnWidth))
                    .InventionREFacilityRegionWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityRegionWidth", DefaultMTColumnWidth))
                    .InventionREFacilityMEBonusWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityMEBonusWidth", DefaultMTColumnWidth))
                    .InventionREFacilityTEBonusWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityTEBonusWidth", DefaultMTColumnWidth))
                    .InventionREFacilityUsageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionREFacilityUsageWidth", DefaultMTColumnWidth))
                    .ManufacturingTeamNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingTeamNameWidth", DefaultMTColumnWidth))
                    .ManufacturingTeamBonusesWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingTeamBonusesWidth", DefaultMTColumnWidth))
                    .ManufacturingTeamUsageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingTeamUsageWidth", DefaultMTColumnWidth))
                    .ManufacturingTeamCostModifierWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ManufacturingTeamCostModifierWidth", DefaultMTColumnWidth))
                    .ComponentTeamNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentTeamNameWidth", DefaultMTColumnWidth))
                    .ComponentTeamBonusesWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentTeamBonusesWidth", DefaultMTColumnWidth))
                    .ComponentTeamUsageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentTeamUsageWidth", DefaultMTColumnWidth))
                    .ComponentTeamCostModifierWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "ComponentTeamCostModifierWidth", DefaultMTColumnWidth))
                    .CopyingTeamNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingTeamNameWidth", DefaultMTColumnWidth))
                    .CopyingTeamBonusesWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingTeamBonusesWidth", DefaultMTColumnWidth))
                    .CopyingTeamUsageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingTeamUsageWidth", DefaultMTColumnWidth))
                    .CopyingTeamCostModifierWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "CopyingTeamCostModifierWidth", DefaultMTColumnWidth))
                    .InventionRETeamNameWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionRETeamNameWidth", DefaultMTColumnWidth))
                    .InventionRETeamBonusesWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionRETeamBonusesWidth", DefaultMTColumnWidth))
                    .InventionRETeamUsageWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionRETeamUsageWidth", DefaultMTColumnWidth))
                    .InventionRETeamCostModifierWidth = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "InventionRETeamCostModifierWidth", DefaultMTColumnWidth))

                    .OrderByColumn = CInt(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeInteger, "ManufacturingTabColumnSettings", "OrderByColumn", DefaultMTOrderByColumn))
                    .OrderType = CStr(GetSettingValue(ManufacturingTabColumnSettingsFileName, SettingTypes.TypeString, "ManufacturingTabColumnSettings", "OrderType", DefaultMTOrderType))

                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultManufacturingTabColumnSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Industry Jobs Column Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultManufacturingTabColumnSettings()
        End Try

        ' Save them locally and then export
        ManufacturingTabColumnSettings = TempSettings

        Return TempSettings

    End Function

    ' Loads the Defaults for the tab
    Public Function SetDefaultManufacturingTabColumnSettings() As ManufacturingTabColumnSettings
        Dim LocalSettings As ManufacturingTabColumnSettings

        With LocalSettings
            .ItemCategory = DefaultMTItemCategory
            .ItemGroup = DefaultMTItemGroup
            .ItemName = DefaultMTItemName
            .Owned = DefaultMTOwned
            .Tech = DefaultMTTech
            .BPME = DefaultMTBPME
            .BPTE = DefaultMTBPTE
            .Inputs = DefaultMTInputs
            .Compared = DefaultMTCompared
            .Runs = DefaultMTRuns
            .ProductionLines = DefaultMTProductionLines
            .LaboratoryLines = DefaultMTLaboratoryLines
            .TotalInventionRECost = DefaultMTTotalInventionRECost
            .TotalCopyCost = DefaultMTTotalCopyCost
            .TotalManufacturingCost = DefaultMTTotalManufacturingCost
            .Taxes = DefaultMTTaxes
            .BrokerFees = DefaultMTBrokerFees
            .BPProductionTime = DefaultMTBPProductionTime
            .TotalProductionTime = DefaultMTTotalProductionTime
            .ItemMarketPrice = DefaultMTItemMarketPrice
            .Profit = DefaultMTProfit
            .ProfitPercentage = DefaultMTProfitPercentage
            .IskperHour = DefaultMTIskperHour
            .SVR = DefaultMTSVR
            .TotalCost = DefaultMTTotalCost
            .BaseJobCost = DefaultMTBaseJobCost
            .ManufacturingJobFee = DefaultMTManufacturingJobFee
            .ManufacturingFacilityName = DefaultMTManufacturingFacilityName
            .ManufacturingFacilitySystem = DefaultMTManufacturingFacilitySystem
            .ManufacturingFacilitySystemIndex = DefaultMTManufacturingFacilitySystemIndex
            .ManufacturingFacilityTax = DefaultMTManufacturingFacilityTax
            .ManufacturingFacilityRegion = DefaultMTManufacturingFacilityRegion
            .ManufacturingFacilityMEBonus = DefaultMTManufacturingFacilityMEBonus
            .ManufacturingFacilityTEBonus = DefaultMTManufacturingFacilityTEBonus
            .ManufacturingFacilityUsage = DefaultMTManufacturingFacilityUsage
            .ComponentFacilityName = DefaultMTComponentFacilityName
            .ComponentFacilitySystem = DefaultMTComponentFacilitySystem
            .ComponentFacilitySystemIndex = DefaultMTComponentFacilitySystemIndex
            .ComponentFacilityTax = DefaultMTComponentFacilityTax
            .ComponentFacilityRegion = DefaultMTComponentFacilityRegion
            .ComponentFacilityMEBonus = DefaultMTComponentFacilityMEBonus
            .ComponentFacilityTEBonus = DefaultMTComponentFacilityTEBonus
            .ComponentFacilityUsage = DefaultMTComponentFacilityUsage
            .CopyingFacilityName = DefaultMTCopyingFacilityName
            .CopyingFacilitySystem = DefaultMTCopyingFacilitySystem
            .CopyingFacilitySystemIndex = DefaultMTCopyingFacilitySystemIndex
            .CopyingFacilityTax = DefaultMTCopyingFacilityTax
            .CopyingFacilityRegion = DefaultMTCopyingFacilityRegion
            .CopyingFacilityMEBonus = DefaultMTCopyingFacilityMEBonus
            .CopyingFacilityTEBonus = DefaultMTCopyingFacilityTEBonus
            .CopyingFacilityUsage = DefaultMTCopyingFacilityUsage
            .InventionREFacilityName = DefaultMTInventionREFacilityName
            .InventionREFacilitySystem = DefaultMTInventionREFacilitySystem
            .InventionREFacilitySystemIndex = DefaultMTInventionREFacilitySystemIndex
            .InventionREFacilityTax = DefaultMTInventionREFacilityTax
            .InventionREFacilityRegion = DefaultMTInventionREFacilityRegion
            .InventionREFacilityMEBonus = DefaultMTInventionREFacilityMEBonus
            .InventionREFacilityTEBonus = DefaultMTInventionREFacilityTEBonus
            .InventionREFacilityUsage = DefaultMTInventionREFacilityUsage
            .ManufacturingTeamName = DefaultMTManufacturingTeamName
            .ManufacturingTeamBonuses = DefaultMTManufacturingTeamBonuses
            .ManufacturingTeamUsage = DefaultMTManufacturingTeamUsage
            .ManufacturingTeamCostModifier = DefaultMTManufacturingTeamCostModifier
            .ComponentTeamName = DefaultMTComponentTeamName
            .ComponentTeamBonuses = DefaultMTComponentTeamBonuses
            .ComponentTeamUsage = DefaultMTComponentTeamUsage
            .ComponentTeamCostModifier = DefaultMTComponentTeamCostModifier
            .CopyingTeamName = DefaultMTCopyingTeamName
            .CopyingTeamBonuses = DefaultMTCopyingTeamBonuses
            .CopyingTeamUsage = DefaultMTCopyingTeamUsage
            .CopyingTeamCostModifier = DefaultMTCopyingTeamCostModifier
            .InventionRETeamName = DefaultMTInventionRETeamName
            .InventionRETeamBonuses = DefaultMTInventionRETeamBonuses
            .InventionRETeamUsage = DefaultMTInventionRETeamUsage
            .InventionRETeamCostModifier = DefaultMTInventionRETeamCostModifier

            .ItemCategoryWidth = DefaultMTColumnWidth
            .ItemGroupWidth = DefaultMTColumnWidth
            .ItemNameWidth = DefaultMTColumnWidth
            .OwnedWidth = DefaultMTColumnWidth
            .TechWidth = DefaultMTColumnWidth
            .BPMEWidth = DefaultMTColumnWidth
            .BPTEWidth = DefaultMTColumnWidth
            .InputsWidth = DefaultMTColumnWidth
            .ComparedWidth = DefaultMTColumnWidth
            .RunsWidth = DefaultMTColumnWidth
            .ProductionLinesWidth = DefaultMTColumnWidth
            .LaboratoryLinesWidth = DefaultMTColumnWidth
            .TotalInventionRECostWidth = DefaultMTColumnWidth
            .TotalCopyCostWidth = DefaultMTColumnWidth
            .TotalManufacturingCostWidth = DefaultMTColumnWidth
            .TaxesWidth = DefaultMTColumnWidth
            .BrokerFeesWidth = DefaultMTColumnWidth
            .BPProductionTimeWidth = DefaultMTColumnWidth
            .TotalProductionTimeWidth = DefaultMTColumnWidth
            .ItemMarketPriceWidth = DefaultMTColumnWidth
            .ProfitWidth = DefaultMTColumnWidth
            .ProfitPercentageWidth = DefaultMTColumnWidth
            .IskperHourWidth = DefaultMTColumnWidth
            .SVRWidth = DefaultMTColumnWidth
            .TotalCostWidth = DefaultMTColumnWidth
            .BaseJobCostWidth = DefaultMTColumnWidth
            .ManufacturingJobFeeWidth = DefaultMTColumnWidth
            .ManufacturingFacilityNameWidth = DefaultMTColumnWidth
            .ManufacturingFacilitySystemWidth = DefaultMTColumnWidth
            .ManufacturingFacilitySystemIndexWidth = DefaultMTColumnWidth
            .ManufacturingFacilityTaxWidth = DefaultMTColumnWidth
            .ManufacturingFacilityRegionWidth = DefaultMTColumnWidth
            .ManufacturingFacilityMEBonusWidth = DefaultMTColumnWidth
            .ManufacturingFacilityTEBonusWidth = DefaultMTColumnWidth
            .ManufacturingFacilityUsageWidth = DefaultMTColumnWidth
            .ComponentFacilityNameWidth = DefaultMTColumnWidth
            .ComponentFacilitySystemWidth = DefaultMTColumnWidth
            .ComponentFacilitySystemIndexWidth = DefaultMTColumnWidth
            .ComponentFacilityTaxWidth = DefaultMTColumnWidth
            .ComponentFacilityRegionWidth = DefaultMTColumnWidth
            .ComponentFacilityMEBonusWidth = DefaultMTColumnWidth
            .ComponentFacilityTEBonusWidth = DefaultMTColumnWidth
            .ComponentFacilityUsageWidth = DefaultMTColumnWidth
            .CopyingFacilityNameWidth = DefaultMTColumnWidth
            .CopyingFacilitySystemWidth = DefaultMTColumnWidth
            .CopyingFacilitySystemIndexWidth = DefaultMTColumnWidth
            .CopyingFacilityTaxWidth = DefaultMTColumnWidth
            .CopyingFacilityRegionWidth = DefaultMTColumnWidth
            .CopyingFacilityMEBonusWidth = DefaultMTColumnWidth
            .CopyingFacilityTEBonusWidth = DefaultMTColumnWidth
            .CopyingFacilityUsageWidth = DefaultMTColumnWidth
            .InventionREFacilityNameWidth = DefaultMTColumnWidth
            .InventionREFacilitySystemWidth = DefaultMTColumnWidth
            .InventionREFacilitySystemIndexWidth = DefaultMTColumnWidth
            .InventionREFacilityTaxWidth = DefaultMTColumnWidth
            .InventionREFacilityRegionWidth = DefaultMTColumnWidth
            .InventionREFacilityMEBonusWidth = DefaultMTColumnWidth
            .InventionREFacilityTEBonusWidth = DefaultMTColumnWidth
            .InventionREFacilityUsageWidth = DefaultMTColumnWidth
            .ManufacturingTeamNameWidth = DefaultMTColumnWidth
            .ManufacturingTeamBonusesWidth = DefaultMTColumnWidth
            .ManufacturingTeamUsageWidth = DefaultMTColumnWidth
            .ManufacturingTeamCostModifierWidth = DefaultMTColumnWidth
            .ComponentTeamNameWidth = DefaultMTColumnWidth
            .ComponentTeamBonusesWidth = DefaultMTColumnWidth
            .ComponentTeamUsageWidth = DefaultMTColumnWidth
            .ComponentTeamCostModifierWidth = DefaultMTColumnWidth
            .CopyingTeamNameWidth = DefaultMTColumnWidth
            .CopyingTeamBonusesWidth = DefaultMTColumnWidth
            .CopyingTeamUsageWidth = DefaultMTColumnWidth
            .CopyingTeamCostModifierWidth = DefaultMTColumnWidth
            .InventionRETeamNameWidth = DefaultMTColumnWidth
            .InventionRETeamBonusesWidth = DefaultMTColumnWidth
            .InventionRETeamUsageWidth = DefaultMTColumnWidth
            .InventionRETeamCostModifierWidth = DefaultMTColumnWidth

            .OrderByColumn = DefaultMTOrderByColumn
            .OrderType = DefaultMTOrderType

        End With

        ' save locally
        ManufacturingTabColumnSettings = LocalSettings
        Return LocalSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveManufacturingTabColumnSettings(SentSettings As ManufacturingTabColumnSettings)
        Dim ManufacturingTabColumnSettingsList(151) As Setting

        Try
            ManufacturingTabColumnSettingsList(0) = New Setting("ItemCategory", CStr(SentSettings.ItemCategory))
            ManufacturingTabColumnSettingsList(1) = New Setting("ItemGroup", CStr(SentSettings.ItemGroup))
            ManufacturingTabColumnSettingsList(2) = New Setting("ItemName", CStr(SentSettings.ItemName))
            ManufacturingTabColumnSettingsList(3) = New Setting("Owned", CStr(SentSettings.Owned))
            ManufacturingTabColumnSettingsList(4) = New Setting("Tech", CStr(SentSettings.Tech))
            ManufacturingTabColumnSettingsList(5) = New Setting("BPME", CStr(SentSettings.BPME))
            ManufacturingTabColumnSettingsList(6) = New Setting("BPTE", CStr(SentSettings.BPTE))
            ManufacturingTabColumnSettingsList(7) = New Setting("Inputs", CStr(SentSettings.Inputs))
            ManufacturingTabColumnSettingsList(8) = New Setting("Compared", CStr(SentSettings.Compared))
            ManufacturingTabColumnSettingsList(9) = New Setting("Runs", CStr(SentSettings.Runs))
            ManufacturingTabColumnSettingsList(10) = New Setting("ProductionLines", CStr(SentSettings.ProductionLines))
            ManufacturingTabColumnSettingsList(11) = New Setting("LaboratoryLines", CStr(SentSettings.LaboratoryLines))
            ManufacturingTabColumnSettingsList(12) = New Setting("TotalInventionRECost", CStr(SentSettings.TotalInventionRECost))
            ManufacturingTabColumnSettingsList(13) = New Setting("TotalCopyCost", CStr(SentSettings.TotalCopyCost))
            ManufacturingTabColumnSettingsList(14) = New Setting("TotalManufacturingCost", CStr(SentSettings.TotalManufacturingCost))
            ManufacturingTabColumnSettingsList(15) = New Setting("Taxes", CStr(SentSettings.Taxes))
            ManufacturingTabColumnSettingsList(16) = New Setting("BrokerFees", CStr(SentSettings.BrokerFees))
            ManufacturingTabColumnSettingsList(17) = New Setting("BPProductionTime", CStr(SentSettings.BPProductionTime))
            ManufacturingTabColumnSettingsList(18) = New Setting("TotalProductionTime", CStr(SentSettings.TotalProductionTime))
            ManufacturingTabColumnSettingsList(19) = New Setting("ItemMarketPrice", CStr(SentSettings.ItemMarketPrice))
            ManufacturingTabColumnSettingsList(20) = New Setting("Profit", CStr(SentSettings.Profit))
            ManufacturingTabColumnSettingsList(21) = New Setting("ProfitPercentage", CStr(SentSettings.ProfitPercentage))
            ManufacturingTabColumnSettingsList(22) = New Setting("IskperHour", CStr(SentSettings.IskperHour))
            ManufacturingTabColumnSettingsList(23) = New Setting("SVR", CStr(SentSettings.SVR))
            ManufacturingTabColumnSettingsList(24) = New Setting("TotalCost", CStr(SentSettings.TotalCost))
            ManufacturingTabColumnSettingsList(25) = New Setting("BaseJobCost", CStr(SentSettings.BaseJobCost))
            ManufacturingTabColumnSettingsList(26) = New Setting("ManufacturingJobFee", CStr(SentSettings.ManufacturingJobFee))
            ManufacturingTabColumnSettingsList(27) = New Setting("ManufacturingFacilityName", CStr(SentSettings.ManufacturingFacilityName))
            ManufacturingTabColumnSettingsList(28) = New Setting("ManufacturingFacilitySystem", CStr(SentSettings.ManufacturingFacilitySystem))
            ManufacturingTabColumnSettingsList(29) = New Setting("ManufacturingFacilitySystemIndex", CStr(SentSettings.ManufacturingFacilitySystemIndex))
            ManufacturingTabColumnSettingsList(30) = New Setting("ManufacturingFacilityTax", CStr(SentSettings.ManufacturingFacilityTax))
            ManufacturingTabColumnSettingsList(31) = New Setting("ManufacturingFacilityRegion", CStr(SentSettings.ManufacturingFacilityRegion))
            ManufacturingTabColumnSettingsList(32) = New Setting("ManufacturingFacilityMEBonus", CStr(SentSettings.ManufacturingFacilityMEBonus))
            ManufacturingTabColumnSettingsList(33) = New Setting("ManufacturingFacilityTEBonus", CStr(SentSettings.ManufacturingFacilityTEBonus))
            ManufacturingTabColumnSettingsList(34) = New Setting("ManufacturingFacilityUsage", CStr(SentSettings.ManufacturingFacilityUsage))
            ManufacturingTabColumnSettingsList(35) = New Setting("ComponentFacilityName", CStr(SentSettings.ComponentFacilityName))
            ManufacturingTabColumnSettingsList(36) = New Setting("ComponentFacilitySystem", CStr(SentSettings.ComponentFacilitySystem))
            ManufacturingTabColumnSettingsList(37) = New Setting("ComponentFacilitySystemIndex", CStr(SentSettings.ComponentFacilitySystemIndex))
            ManufacturingTabColumnSettingsList(38) = New Setting("ComponentFacilityTax", CStr(SentSettings.ComponentFacilityTax))
            ManufacturingTabColumnSettingsList(39) = New Setting("ComponentFacilityRegion", CStr(SentSettings.ComponentFacilityRegion))
            ManufacturingTabColumnSettingsList(40) = New Setting("ComponentFacilityMEBonus", CStr(SentSettings.ComponentFacilityMEBonus))
            ManufacturingTabColumnSettingsList(41) = New Setting("ComponentFacilityTEBonus", CStr(SentSettings.ComponentFacilityTEBonus))
            ManufacturingTabColumnSettingsList(42) = New Setting("ComponentFacilityUsage", CStr(SentSettings.ComponentFacilityUsage))
            ManufacturingTabColumnSettingsList(43) = New Setting("CopyingFacilityName", CStr(SentSettings.CopyingFacilityName))
            ManufacturingTabColumnSettingsList(44) = New Setting("CopyingFacilitySystem", CStr(SentSettings.CopyingFacilitySystem))
            ManufacturingTabColumnSettingsList(45) = New Setting("CopyingFacilitySystemIndex", CStr(SentSettings.CopyingFacilitySystemIndex))
            ManufacturingTabColumnSettingsList(46) = New Setting("CopyingFacilityTax", CStr(SentSettings.CopyingFacilityTax))
            ManufacturingTabColumnSettingsList(47) = New Setting("CopyingFacilityRegion", CStr(SentSettings.CopyingFacilityRegion))
            ManufacturingTabColumnSettingsList(48) = New Setting("CopyingFacilityMEBonus", CStr(SentSettings.CopyingFacilityMEBonus))
            ManufacturingTabColumnSettingsList(49) = New Setting("CopyingFacilityTEBonus", CStr(SentSettings.CopyingFacilityTEBonus))
            ManufacturingTabColumnSettingsList(50) = New Setting("CopyingFacilityUsage", CStr(SentSettings.CopyingFacilityUsage))
            ManufacturingTabColumnSettingsList(51) = New Setting("InventionREFacilityName", CStr(SentSettings.InventionREFacilityName))
            ManufacturingTabColumnSettingsList(52) = New Setting("InventionREFacilitySystem", CStr(SentSettings.InventionREFacilitySystem))
            ManufacturingTabColumnSettingsList(53) = New Setting("InventionREFacilitySystemIndex", CStr(SentSettings.InventionREFacilitySystemIndex))
            ManufacturingTabColumnSettingsList(54) = New Setting("InventionREFacilityTax", CStr(SentSettings.InventionREFacilityTax))
            ManufacturingTabColumnSettingsList(55) = New Setting("InventionREFacilityRegion", CStr(SentSettings.InventionREFacilityRegion))
            ManufacturingTabColumnSettingsList(56) = New Setting("InventionREFacilityMEBonus", CStr(SentSettings.InventionREFacilityMEBonus))
            ManufacturingTabColumnSettingsList(57) = New Setting("InventionREFacilityTEBonus", CStr(SentSettings.InventionREFacilityTEBonus))
            ManufacturingTabColumnSettingsList(58) = New Setting("InventionREFacilityUsage", CStr(SentSettings.InventionREFacilityUsage))
            ManufacturingTabColumnSettingsList(59) = New Setting("ManufacturingTeamName", CStr(SentSettings.ManufacturingTeamName))
            ManufacturingTabColumnSettingsList(60) = New Setting("ManufacturingTeamBonuses", CStr(SentSettings.ManufacturingTeamBonuses))
            ManufacturingTabColumnSettingsList(61) = New Setting("ManufacturingTeamUsage", CStr(SentSettings.ManufacturingTeamUsage))
            ManufacturingTabColumnSettingsList(62) = New Setting("ManufacturingTeamCostModifier", CStr(SentSettings.ManufacturingTeamCostModifier))
            ManufacturingTabColumnSettingsList(63) = New Setting("ComponentTeamName", CStr(SentSettings.ComponentTeamName))
            ManufacturingTabColumnSettingsList(64) = New Setting("ComponentTeamBonuses", CStr(SentSettings.ComponentTeamBonuses))
            ManufacturingTabColumnSettingsList(65) = New Setting("ComponentTeamUsage", CStr(SentSettings.ComponentTeamUsage))
            ManufacturingTabColumnSettingsList(66) = New Setting("ComponentTeamCostModifier", CStr(SentSettings.ComponentTeamCostModifier))
            ManufacturingTabColumnSettingsList(67) = New Setting("CopyingTeamName", CStr(SentSettings.CopyingTeamName))
            ManufacturingTabColumnSettingsList(68) = New Setting("CopyingTeamBonuses", CStr(SentSettings.CopyingTeamBonuses))
            ManufacturingTabColumnSettingsList(69) = New Setting("CopyingTeamUsage", CStr(SentSettings.CopyingTeamUsage))
            ManufacturingTabColumnSettingsList(70) = New Setting("CopyingTeamCostModifier", CStr(SentSettings.CopyingTeamCostModifier))
            ManufacturingTabColumnSettingsList(71) = New Setting("InventionRETeamName", CStr(SentSettings.InventionRETeamName))
            ManufacturingTabColumnSettingsList(72) = New Setting("InventionRETeamBonuses", CStr(SentSettings.InventionRETeamBonuses))
            ManufacturingTabColumnSettingsList(73) = New Setting("InventionRETeamUsage", CStr(SentSettings.InventionRETeamUsage))
            ManufacturingTabColumnSettingsList(74) = New Setting("InventionRETeamCostModifier", CStr(SentSettings.InventionRETeamCostModifier))

            ManufacturingTabColumnSettingsList(75) = New Setting("ItemCategoryWidth", CStr(SentSettings.ItemCategoryWidth))
            ManufacturingTabColumnSettingsList(76) = New Setting("ItemGroupWidth", CStr(SentSettings.ItemGroupWidth))
            ManufacturingTabColumnSettingsList(77) = New Setting("ItemNameWidth", CStr(SentSettings.ItemNameWidth))
            ManufacturingTabColumnSettingsList(78) = New Setting("OwnedWidth", CStr(SentSettings.OwnedWidth))
            ManufacturingTabColumnSettingsList(79) = New Setting("TechWidth", CStr(SentSettings.TechWidth))
            ManufacturingTabColumnSettingsList(80) = New Setting("BPMEWidth", CStr(SentSettings.BPMEWidth))
            ManufacturingTabColumnSettingsList(81) = New Setting("BPTEWidth", CStr(SentSettings.BPTEWidth))
            ManufacturingTabColumnSettingsList(82) = New Setting("InputsWidth", CStr(SentSettings.InputsWidth))
            ManufacturingTabColumnSettingsList(83) = New Setting("ComparedWidth", CStr(SentSettings.ComparedWidth))
            ManufacturingTabColumnSettingsList(84) = New Setting("RunsWidth", CStr(SentSettings.RunsWidth))
            ManufacturingTabColumnSettingsList(85) = New Setting("ProductionLinesWidth", CStr(SentSettings.ProductionLinesWidth))
            ManufacturingTabColumnSettingsList(86) = New Setting("LaboratoryLinesWidth", CStr(SentSettings.LaboratoryLinesWidth))
            ManufacturingTabColumnSettingsList(87) = New Setting("TotalInventionRECostWidth", CStr(SentSettings.TotalInventionRECostWidth))
            ManufacturingTabColumnSettingsList(88) = New Setting("TotalCopyCostWidth", CStr(SentSettings.TotalCopyCostWidth))
            ManufacturingTabColumnSettingsList(89) = New Setting("TotalManufacturingCostWidth", CStr(SentSettings.TotalManufacturingCostWidth))
            ManufacturingTabColumnSettingsList(90) = New Setting("TaxesWidth", CStr(SentSettings.TaxesWidth))
            ManufacturingTabColumnSettingsList(91) = New Setting("BrokerFeesWidth", CStr(SentSettings.BrokerFeesWidth))
            ManufacturingTabColumnSettingsList(92) = New Setting("BPProductionTimeWidth", CStr(SentSettings.BPProductionTimeWidth))
            ManufacturingTabColumnSettingsList(93) = New Setting("TotalProductionTimeWidth", CStr(SentSettings.TotalProductionTimeWidth))
            ManufacturingTabColumnSettingsList(94) = New Setting("ItemMarketPriceWidth", CStr(SentSettings.ItemMarketPriceWidth))
            ManufacturingTabColumnSettingsList(95) = New Setting("ProfitWidth", CStr(SentSettings.ProfitWidth))
            ManufacturingTabColumnSettingsList(96) = New Setting("ProfitPercentageWidth", CStr(SentSettings.ProfitPercentageWidth))
            ManufacturingTabColumnSettingsList(97) = New Setting("IskperHourWidth", CStr(SentSettings.IskperHourWidth))
            ManufacturingTabColumnSettingsList(98) = New Setting("SVRWidth", CStr(SentSettings.SVRWidth))
            ManufacturingTabColumnSettingsList(99) = New Setting("TotalCostWidth", CStr(SentSettings.TotalCostWidth))
            ManufacturingTabColumnSettingsList(100) = New Setting("BaseJobCostWidth", CStr(SentSettings.BaseJobCostWidth))
            ManufacturingTabColumnSettingsList(101) = New Setting("ManufacturingJobFeeWidth", CStr(SentSettings.ManufacturingJobFeeWidth))
            ManufacturingTabColumnSettingsList(102) = New Setting("ManufacturingFacilityNameWidth", CStr(SentSettings.ManufacturingFacilityNameWidth))
            ManufacturingTabColumnSettingsList(103) = New Setting("ManufacturingFacilitySystemWidth", CStr(SentSettings.ManufacturingFacilitySystemWidth))
            ManufacturingTabColumnSettingsList(104) = New Setting("ManufacturingFacilitySystemIndexWidth", CStr(SentSettings.ManufacturingFacilitySystemIndexWidth))
            ManufacturingTabColumnSettingsList(105) = New Setting("ManufacturingFacilityTaxWidth", CStr(SentSettings.ManufacturingFacilityTaxWidth))
            ManufacturingTabColumnSettingsList(106) = New Setting("ManufacturingFacilityRegionWidth", CStr(SentSettings.ManufacturingFacilityRegionWidth))
            ManufacturingTabColumnSettingsList(107) = New Setting("ManufacturingFacilityMEBonusWidth", CStr(SentSettings.ManufacturingFacilityMEBonusWidth))
            ManufacturingTabColumnSettingsList(108) = New Setting("ManufacturingFacilityTEBonusWidth", CStr(SentSettings.ManufacturingFacilityTEBonusWidth))
            ManufacturingTabColumnSettingsList(109) = New Setting("ManufacturingFacilityUsageWidth", CStr(SentSettings.ManufacturingFacilityUsageWidth))
            ManufacturingTabColumnSettingsList(110) = New Setting("ComponentFacilityNameWidth", CStr(SentSettings.ComponentFacilityNameWidth))
            ManufacturingTabColumnSettingsList(111) = New Setting("ComponentFacilitySystemWidth", CStr(SentSettings.ComponentFacilitySystemWidth))
            ManufacturingTabColumnSettingsList(112) = New Setting("ComponentFacilitySystemIndexWidth", CStr(SentSettings.ComponentFacilitySystemIndexWidth))
            ManufacturingTabColumnSettingsList(113) = New Setting("ComponentFacilityTaxWidth", CStr(SentSettings.ComponentFacilityTaxWidth))
            ManufacturingTabColumnSettingsList(114) = New Setting("ComponentFacilityRegionWidth", CStr(SentSettings.ComponentFacilityRegionWidth))
            ManufacturingTabColumnSettingsList(115) = New Setting("ComponentFacilityMEBonusWidth", CStr(SentSettings.ComponentFacilityMEBonusWidth))
            ManufacturingTabColumnSettingsList(116) = New Setting("ComponentFacilityTEBonusWidth", CStr(SentSettings.ComponentFacilityTEBonusWidth))
            ManufacturingTabColumnSettingsList(117) = New Setting("ComponentFacilityUsageWidth", CStr(SentSettings.ComponentFacilityUsageWidth))
            ManufacturingTabColumnSettingsList(118) = New Setting("CopyingFacilityNameWidth", CStr(SentSettings.CopyingFacilityNameWidth))
            ManufacturingTabColumnSettingsList(119) = New Setting("CopyingFacilitySystemWidth", CStr(SentSettings.CopyingFacilitySystemWidth))
            ManufacturingTabColumnSettingsList(120) = New Setting("CopyingFacilitySystemIndexWidth", CStr(SentSettings.CopyingFacilitySystemIndexWidth))
            ManufacturingTabColumnSettingsList(121) = New Setting("CopyingFacilityTaxWidth", CStr(SentSettings.CopyingFacilityTaxWidth))
            ManufacturingTabColumnSettingsList(122) = New Setting("CopyingFacilityRegionWidth", CStr(SentSettings.CopyingFacilityRegionWidth))
            ManufacturingTabColumnSettingsList(123) = New Setting("CopyingFacilityMEBonusWidth", CStr(SentSettings.CopyingFacilityMEBonusWidth))
            ManufacturingTabColumnSettingsList(124) = New Setting("CopyingFacilityTEBonusWidth", CStr(SentSettings.CopyingFacilityTEBonusWidth))
            ManufacturingTabColumnSettingsList(125) = New Setting("CopyingFacilityUsageWidth", CStr(SentSettings.CopyingFacilityUsageWidth))
            ManufacturingTabColumnSettingsList(126) = New Setting("InventionREFacilityNameWidth", CStr(SentSettings.InventionREFacilityNameWidth))
            ManufacturingTabColumnSettingsList(127) = New Setting("InventionREFacilitySystemWidth", CStr(SentSettings.InventionREFacilitySystemWidth))
            ManufacturingTabColumnSettingsList(128) = New Setting("InventionREFacilitySystemIndexWidth", CStr(SentSettings.InventionREFacilitySystemIndexWidth))
            ManufacturingTabColumnSettingsList(129) = New Setting("InventionREFacilityTaxWidth", CStr(SentSettings.InventionREFacilityTaxWidth))
            ManufacturingTabColumnSettingsList(130) = New Setting("InventionREFacilityRegionWidth", CStr(SentSettings.InventionREFacilityRegionWidth))
            ManufacturingTabColumnSettingsList(131) = New Setting("InventionREFacilityMEBonusWidth", CStr(SentSettings.InventionREFacilityMEBonusWidth))
            ManufacturingTabColumnSettingsList(132) = New Setting("InventionREFacilityTEBonusWidth", CStr(SentSettings.InventionREFacilityTEBonusWidth))
            ManufacturingTabColumnSettingsList(133) = New Setting("InventionREFacilityUsageWidth", CStr(SentSettings.InventionREFacilityUsageWidth))
            ManufacturingTabColumnSettingsList(134) = New Setting("ManufacturingTeamNameWidth", CStr(SentSettings.ManufacturingTeamNameWidth))
            ManufacturingTabColumnSettingsList(135) = New Setting("ManufacturingTeamBonusesWidth", CStr(SentSettings.ManufacturingTeamBonusesWidth))
            ManufacturingTabColumnSettingsList(136) = New Setting("ManufacturingTeamUsageWidth", CStr(SentSettings.ManufacturingTeamUsageWidth))
            ManufacturingTabColumnSettingsList(137) = New Setting("ManufacturingTeamCostModifierWidth", CStr(SentSettings.ManufacturingTeamCostModifierWidth))
            ManufacturingTabColumnSettingsList(138) = New Setting("ComponentTeamNameWidth", CStr(SentSettings.ComponentTeamNameWidth))
            ManufacturingTabColumnSettingsList(139) = New Setting("ComponentTeamBonusesWidth", CStr(SentSettings.ComponentTeamBonusesWidth))
            ManufacturingTabColumnSettingsList(140) = New Setting("ComponentTeamUsageWidth", CStr(SentSettings.ComponentTeamUsageWidth))
            ManufacturingTabColumnSettingsList(141) = New Setting("ComponentTeamCostModifierWidth", CStr(SentSettings.ComponentTeamCostModifierWidth))
            ManufacturingTabColumnSettingsList(142) = New Setting("CopyingTeamNameWidth", CStr(SentSettings.CopyingTeamNameWidth))
            ManufacturingTabColumnSettingsList(143) = New Setting("CopyingTeamBonusesWidth", CStr(SentSettings.CopyingTeamBonusesWidth))
            ManufacturingTabColumnSettingsList(144) = New Setting("CopyingTeamUsageWidth", CStr(SentSettings.CopyingTeamUsageWidth))
            ManufacturingTabColumnSettingsList(145) = New Setting("CopyingTeamCostModifierWidth", CStr(SentSettings.CopyingTeamCostModifierWidth))
            ManufacturingTabColumnSettingsList(146) = New Setting("InventionRETeamNameWidth", CStr(SentSettings.InventionRETeamNameWidth))
            ManufacturingTabColumnSettingsList(147) = New Setting("InventionRETeamBonusesWidth", CStr(SentSettings.InventionRETeamBonusesWidth))
            ManufacturingTabColumnSettingsList(148) = New Setting("InventionRETeamUsageWidth", CStr(SentSettings.InventionRETeamUsageWidth))
            ManufacturingTabColumnSettingsList(149) = New Setting("InventionRETeamCostModifierWidth", CStr(SentSettings.InventionRETeamCostModifierWidth))

            ManufacturingTabColumnSettingsList(150) = New Setting("OrderMTByColumn", CStr(SentSettings.OrderByColumn))
            ManufacturingTabColumnSettingsList(151) = New Setting("OrderMTType", CStr(SentSettings.OrderType))

            Call WriteSettingsToFile(ManufacturingTabColumnSettingsFileName, ManufacturingTabColumnSettingsList, "ManufacturingTabColumnSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Industry Jobs Column Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetManufacturingTabColumnSettings() As ManufacturingTabColumnSettings
        Return ManufacturingTabColumnSettings
    End Function

#End Region

#Region "Industry Belt Flip"

    ' Loads the tab settings
    Public Function LoadIndustryFlipBeltColumnSettings() As IndustryFlipBeltSettings
        Dim TempSettings As IndustryFlipBeltSettings = Nothing

        Try
            If File.Exists(SettingsFolder & IndustryFlipBeltSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .CycleTime = CDbl(GetSettingValue(IndustryFlipBeltSettingsFileName, SettingTypes.TypeDouble, "IndustryFlipBeltSettings", "CycleTime", DefaultCycleTime))
                    .m3perCycle = CDbl(GetSettingValue(IndustryFlipBeltSettingsFileName, SettingTypes.TypeDouble, "IndustryFlipBeltSettings", "m3perCycle", Defaultm3perCycle))
                    .NumMiners = CInt(GetSettingValue(IndustryFlipBeltSettingsFileName, SettingTypes.TypeInteger, "IndustryFlipBeltSettings", "NumMiners", DefaultNumMiners))
                    .CompressOre = CBool(GetSettingValue(IndustryFlipBeltSettingsFileName, SettingTypes.TypeBoolean, "IndustryFlipBeltSettings", "CompressOre", DefaultCompressOre))
                    .IPHperMiner = CBool(GetSettingValue(IndustryFlipBeltSettingsFileName, SettingTypes.TypeBoolean, "IndustryFlipBeltSettings", "IPHperMiner", DefaultIPHperMiner))
                    .IncludeBrokerFees = CBool(GetSettingValue(IndustryFlipBeltSettingsFileName, SettingTypes.TypeBoolean, "IndustryFlipBeltSettings", "IncludeBrokerFees", DefaultIncludeBrokerFees))
                    .IncludeTaxes = CBool(GetSettingValue(IndustryFlipBeltSettingsFileName, SettingTypes.TypeBoolean, "IndustryFlipBeltSettings", "IncludeTaxes", DefaultIncludeTaxes))
                    .TrueSec = CStr(GetSettingValue(IndustryFlipBeltSettingsFileName, SettingTypes.TypeString, "IndustryFlipBeltSettings", "TrueSec", DefaultTruesec))
                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultIndustryFlipBeltSettings()
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Industry Flip Belt Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultIndustryFlipBeltSettings()
        End Try

        ' Save them locally and then export
        IndustryFlipBeltsSettings = TempSettings

        Return TempSettings

    End Function

    ' Loads the Defaults for the tab
    Public Function SetDefaultIndustryFlipBeltSettings() As IndustryFlipBeltSettings
        Dim LocalSettings As IndustryFlipBeltSettings

        With LocalSettings
            .CycleTime = DefaultCycleTime
            .m3perCycle = Defaultm3perCycle
            .NumMiners = DefaultNumMiners
            .CompressOre = DefaultCompressOre
            .IPHperMiner = DefaultIPHperMiner
            .IncludeBrokerFees = DefaultIncludeBrokerFees
            .IncludeTaxes = DefaultIncludeTaxes
            .TrueSec = DefaultTruesec
        End With

        ' Save locally
        IndustryFlipBeltsSettings = LocalSettings
        Return LocalSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveIndustryFlipBeltSettings(SentSettings As IndustryFlipBeltSettings)
        Dim IndustryFlipBeltSettingsList(7) As Setting

        Try
            IndustryFlipBeltSettingsList(0) = New Setting("CycleTime", CStr(SentSettings.CycleTime))
            IndustryFlipBeltSettingsList(1) = New Setting("m3perCycle", CStr(SentSettings.m3perCycle))
            IndustryFlipBeltSettingsList(2) = New Setting("NumMiners", CStr(SentSettings.NumMiners))
            IndustryFlipBeltSettingsList(3) = New Setting("CompressOre", CStr(SentSettings.CompressOre))
            IndustryFlipBeltSettingsList(4) = New Setting("IPHperMiner", CStr(SentSettings.IPHperMiner))
            IndustryFlipBeltSettingsList(5) = New Setting("IncludeBrokerFees", CStr(SentSettings.IncludeBrokerFees))
            IndustryFlipBeltSettingsList(6) = New Setting("IncludeTaxes", CStr(SentSettings.IncludeTaxes))
            IndustryFlipBeltSettingsList(7) = New Setting("TrueSec", CStr(SentSettings.TrueSec))

            Call WriteSettingsToFile(IndustryFlipBeltSettingsFileName, IndustryFlipBeltSettingsList, "IndustryFlipBeltSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Industry Flip Belt Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetIndustryFlipBeltSettings() As IndustryFlipBeltSettings
        Return IndustryFlipBeltsSettings
    End Function

#End Region

#Region "Industry Belt Ore Checks"

    ' Loads the tab settings
    Public Function LoadIndustryBeltOreChecksSettings(Belt As BeltType) As IndustryBeltOreChecks
        Dim TempSettings As IndustryBeltOreChecks = Nothing
        Dim IndustryBeltOreChecksFileName As String = ""

        Select Case Belt
            Case BeltType.Small
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName1
            Case BeltType.Moderate
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName2
            Case BeltType.Large
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName3
            Case BeltType.ExtraLarge
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName4
            Case BeltType.Giant
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName5
        End Select

        Try
            If File.Exists(SettingsFolder & IndustryBeltOreChecksFileName) Then
                'Get the settings
                With TempSettings
                    .Plagioclase = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Plagioclase", DefaultPlagioclase))
                    .Spodumain = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Spodumain", DefaultSpodumain))
                    .Kernite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Kernite", DefaultKernite))
                    .Hedbergite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Hedbergite", DefaultHedbergite))
                    .Arkonor = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Arkonor", DefaultArkonor))
                    .Bistot = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Bistot", DefaultBistot))
                    .Pyroxeres = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Pyroxeres", DefaultPyroxeres))
                    .Crokite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Crokite", DefaultCrokite))
                    .Jaspet = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Jaspet", DefaultJaspet))
                    .Omber = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Omber", DefaultOmber))
                    .Scordite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Scordite", DefaultScordite))
                    .Gneiss = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Gneiss", DefaultGneiss))
                    .Veldspar = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Veldspar", DefaultVeldspar))
                    .Hemorphite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Hemorphite", DefaultHemorphite))
                    .DarkOchre = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "DarkOchre", DefaultDarkOchre))
                    .Mercoxit = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "Mercoxit", DefaultMercoxit))
                    .CrimsonArkonor = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "CrimsonArkonor", DefaultCrimsonArkonor))
                    .PrimeArkonor = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "PrimeArkonor", DefaultPrimeArkonor))
                    .TriclinicBistot = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "TriclinicBistot", DefaultTriclinicBistot))
                    .MonoclinicBistot = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "MonoclinicBistot", DefaultMonoclinicBistot))
                    .SharpCrokite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "SharpCrokite", DefaultSharpCrokite))
                    .CrystallineCrokite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "CrystallineCrokite", DefaultCrystallineCrokite))
                    .OnyxOchre = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "OnyxOchre", DefaultOnyxOchre))
                    .ObsidianOchre = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "ObsidianOchre", DefaultObsidianOchre))
                    .VitricHedbergite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "VitricHedbergite", DefaultVitricHedbergite))
                    .GlazedHedbergite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "GlazedHedbergite", DefaultGlazedHedbergite))
                    .VividHemorphite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "VividHemorphite", DefaultVividHemorphite))
                    .RadiantHemorphite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "RadiantHemorphite", DefaultRadiantHemorphite))
                    .PureJaspet = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "PureJaspet", DefaultPureJaspet))
                    .PristineJaspet = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "PristineJaspet", DefaultPristineJaspet))
                    .LuminousKernite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "LuminousKernite", DefaultLuminousKernite))
                    .FieryKernite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "FieryKernite", DefaultFieryKernite))
                    .AzurePlagioclase = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "AzurePlagioclase", DefaultAzurePlagioclase))
                    .RichPlagioclase = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "RichPlagioclase", DefaultRichPlagioclase))
                    .SolidPyroxeres = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "SolidPyroxeres", DefaultSolidPyroxeres))
                    .ViscousPyroxeres = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "ViscousPyroxeres", DefaultViscousPyroxeres))
                    .CondensedScordite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "CondensedScordite", DefaultCondensedScordite))
                    .MassiveScordite = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "MassiveScordite", DefaultMassiveScordite))
                    .BrightSpodumain = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "BrightSpodumain", DefaultBrightSpodumain))
                    .GleamingSpodumain = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "GleamingSpodumain", DefaultGleamingSpodumain))
                    .ConcentratedVeldspar = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "ConcentratedVeldspar", DefaultConcentratedVeldspar))
                    .DenseVeldspar = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "DenseVeldspar", DefaultDenseVeldspar))
                    .IridescentGneiss = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "IridescentGneiss", DefaultIridescentGneiss))
                    .PrismaticGneiss = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "PrismaticGneiss", DefaultPrismaticGneiss))
                    .SilveryOmber = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "SilveryOmber", DefaultSilveryOmber))
                    .GoldenOmber = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "GoldenOmber", DefaultGoldenOmber))
                    .MagmaMercoxit = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "MagmaMercoxit", DefaultMagmaMercoxit))
                    .VitreousMercoxit = CBool(GetSettingValue(IndustryBeltOreChecksFileName, SettingTypes.TypeBoolean, "IndustryBeltOreChecksSettings", "VitreousMercoxit", DefaultVitreousMercoxit))
                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultIndustryBeltOreChecksSettings(Belt)
            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Industry Flip Belt Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultIndustryBeltOreChecksSettings(Belt)
        End Try

        ' Save them locally and then export
        Select Case Belt
            Case BeltType.Small
                IndustryBeltOreChecksSettings1 = TempSettings
            Case BeltType.Moderate
                IndustryBeltOreChecksSettings2 = TempSettings
            Case BeltType.Large
                IndustryBeltOreChecksSettings3 = TempSettings
            Case BeltType.ExtraLarge
                IndustryBeltOreChecksSettings4 = TempSettings
            Case BeltType.Giant
                IndustryBeltOreChecksSettings5 = TempSettings
        End Select

        Return TempSettings

    End Function

    ' Loads the Defaults for the tab
    Public Function SetDefaultIndustryBeltOreChecksSettings(Belt As BeltType) As IndustryBeltOreChecks
        Dim LocalSettings As IndustryBeltOreChecks

        With LocalSettings
            .Plagioclase = DefaultPlagioclase
            .Spodumain = DefaultSpodumain
            .Kernite = DefaultKernite
            .Hedbergite = DefaultHedbergite
            .Arkonor = DefaultArkonor
            .Bistot = DefaultBistot
            .Pyroxeres = DefaultPyroxeres
            .Crokite = DefaultCrokite
            .Jaspet = DefaultJaspet
            .Omber = DefaultOmber
            .Scordite = DefaultScordite
            .Gneiss = DefaultGneiss
            .Veldspar = DefaultVeldspar
            .Hemorphite = DefaultHemorphite
            .DarkOchre = DefaultDarkOchre
            .Mercoxit = DefaultMercoxit
            .CrimsonArkonor = DefaultCrimsonArkonor
            .PrimeArkonor = DefaultPrimeArkonor
            .TriclinicBistot = DefaultTriclinicBistot
            .MonoclinicBistot = DefaultMonoclinicBistot
            .SharpCrokite = DefaultSharpCrokite
            .CrystallineCrokite = DefaultCrystallineCrokite
            .OnyxOchre = DefaultOnyxOchre
            .ObsidianOchre = DefaultObsidianOchre
            .VitricHedbergite = DefaultVitricHedbergite
            .GlazedHedbergite = DefaultGlazedHedbergite
            .VividHemorphite = DefaultVividHemorphite
            .RadiantHemorphite = DefaultRadiantHemorphite
            .PureJaspet = DefaultPureJaspet
            .PristineJaspet = DefaultPristineJaspet
            .LuminousKernite = DefaultLuminousKernite
            .FieryKernite = DefaultFieryKernite
            .AzurePlagioclase = DefaultAzurePlagioclase
            .RichPlagioclase = DefaultRichPlagioclase
            .SolidPyroxeres = DefaultSolidPyroxeres
            .ViscousPyroxeres = DefaultViscousPyroxeres
            .CondensedScordite = DefaultCondensedScordite
            .MassiveScordite = DefaultMassiveScordite
            .BrightSpodumain = DefaultBrightSpodumain
            .GleamingSpodumain = DefaultGleamingSpodumain
            .ConcentratedVeldspar = DefaultConcentratedVeldspar
            .DenseVeldspar = DefaultDenseVeldspar
            .IridescentGneiss = DefaultIridescentGneiss
            .PrismaticGneiss = DefaultPrismaticGneiss
            .SilveryOmber = DefaultSilveryOmber
            .GoldenOmber = DefaultGoldenOmber
            .MagmaMercoxit = DefaultMagmaMercoxit
            .VitreousMercoxit = DefaultVitreousMercoxit
        End With

        ' Save locally
        ' Save them locally and then export
        Select Case Belt
            Case BeltType.Small
                IndustryBeltOreChecksSettings1 = LocalSettings
            Case BeltType.Moderate
                IndustryBeltOreChecksSettings2 = LocalSettings
            Case BeltType.Large
                IndustryBeltOreChecksSettings3 = LocalSettings
            Case BeltType.ExtraLarge
                IndustryBeltOreChecksSettings4 = LocalSettings
            Case BeltType.Giant
                IndustryBeltOreChecksSettings5 = LocalSettings
        End Select

        Return LocalSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveIndustryBeltOreChecksSettings(SentSettings As IndustryBeltOreChecks, Belt As BeltType)
        Dim IndustryBeltOreChecksList(47) As Setting
        Dim IndustryBeltOreChecksFileName As String = ""

        Select Case Belt
            Case BeltType.Small
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName1
            Case BeltType.Moderate
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName2
            Case BeltType.Large
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName3
            Case BeltType.ExtraLarge
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName4
            Case BeltType.Giant
                IndustryBeltOreChecksFileName = IndustryBeltOreChecksFileName5
        End Select

        Try
            IndustryBeltOreChecksList(0) = New Setting("Plagioclase", CStr(SentSettings.Plagioclase))
            IndustryBeltOreChecksList(1) = New Setting("Spodumain", CStr(SentSettings.Spodumain))
            IndustryBeltOreChecksList(2) = New Setting("Kernite", CStr(SentSettings.Kernite))
            IndustryBeltOreChecksList(3) = New Setting("Hedbergite", CStr(SentSettings.Hedbergite))
            IndustryBeltOreChecksList(4) = New Setting("Arkonor", CStr(SentSettings.Arkonor))
            IndustryBeltOreChecksList(5) = New Setting("Bistot", CStr(SentSettings.Bistot))
            IndustryBeltOreChecksList(6) = New Setting("Pyroxeres", CStr(SentSettings.Pyroxeres))
            IndustryBeltOreChecksList(7) = New Setting("Crokite", CStr(SentSettings.Crokite))
            IndustryBeltOreChecksList(8) = New Setting("Jaspet", CStr(SentSettings.Jaspet))
            IndustryBeltOreChecksList(9) = New Setting("Omber", CStr(SentSettings.Omber))
            IndustryBeltOreChecksList(10) = New Setting("Scordite", CStr(SentSettings.Scordite))
            IndustryBeltOreChecksList(11) = New Setting("Gneiss", CStr(SentSettings.Gneiss))
            IndustryBeltOreChecksList(12) = New Setting("Veldspar", CStr(SentSettings.Veldspar))
            IndustryBeltOreChecksList(13) = New Setting("Hemorphite", CStr(SentSettings.Hemorphite))
            IndustryBeltOreChecksList(14) = New Setting("DarkOchre", CStr(SentSettings.DarkOchre))
            IndustryBeltOreChecksList(15) = New Setting("Mercoxit", CStr(SentSettings.Mercoxit))
            IndustryBeltOreChecksList(16) = New Setting("CrimsonArkonor", CStr(SentSettings.CrimsonArkonor))
            IndustryBeltOreChecksList(17) = New Setting("PrimeArkonor", CStr(SentSettings.PrimeArkonor))
            IndustryBeltOreChecksList(18) = New Setting("TriclinicBistot", CStr(SentSettings.TriclinicBistot))
            IndustryBeltOreChecksList(19) = New Setting("MonoclinicBistot", CStr(SentSettings.MonoclinicBistot))
            IndustryBeltOreChecksList(20) = New Setting("SharpCrokite", CStr(SentSettings.SharpCrokite))
            IndustryBeltOreChecksList(21) = New Setting("CrystallineCrokite", CStr(SentSettings.CrystallineCrokite))
            IndustryBeltOreChecksList(22) = New Setting("OnyxOchre", CStr(SentSettings.OnyxOchre))
            IndustryBeltOreChecksList(23) = New Setting("ObsidianOchre", CStr(SentSettings.ObsidianOchre))
            IndustryBeltOreChecksList(24) = New Setting("VitricHedbergite", CStr(SentSettings.VitricHedbergite))
            IndustryBeltOreChecksList(25) = New Setting("GlazedHedbergite", CStr(SentSettings.GlazedHedbergite))
            IndustryBeltOreChecksList(26) = New Setting("VividHemorphite", CStr(SentSettings.VividHemorphite))
            IndustryBeltOreChecksList(27) = New Setting("RadiantHemorphite", CStr(SentSettings.RadiantHemorphite))
            IndustryBeltOreChecksList(28) = New Setting("PureJaspet", CStr(SentSettings.PureJaspet))
            IndustryBeltOreChecksList(29) = New Setting("PristineJaspet", CStr(SentSettings.PristineJaspet))
            IndustryBeltOreChecksList(30) = New Setting("LuminousKernite", CStr(SentSettings.LuminousKernite))
            IndustryBeltOreChecksList(31) = New Setting("FieryKernite", CStr(SentSettings.FieryKernite))
            IndustryBeltOreChecksList(32) = New Setting("AzurePlagioclase", CStr(SentSettings.AzurePlagioclase))
            IndustryBeltOreChecksList(33) = New Setting("RichPlagioclase", CStr(SentSettings.RichPlagioclase))
            IndustryBeltOreChecksList(34) = New Setting("SolidPyroxeres", CStr(SentSettings.SolidPyroxeres))
            IndustryBeltOreChecksList(35) = New Setting("ViscousPyroxeres", CStr(SentSettings.ViscousPyroxeres))
            IndustryBeltOreChecksList(36) = New Setting("CondensedScordite", CStr(SentSettings.CondensedScordite))
            IndustryBeltOreChecksList(37) = New Setting("MassiveScordite", CStr(SentSettings.MassiveScordite))
            IndustryBeltOreChecksList(38) = New Setting("BrightSpodumain", CStr(SentSettings.BrightSpodumain))
            IndustryBeltOreChecksList(39) = New Setting("GleamingSpodumain", CStr(SentSettings.GleamingSpodumain))
            IndustryBeltOreChecksList(40) = New Setting("ConcentratedVeldspar", CStr(SentSettings.ConcentratedVeldspar))
            IndustryBeltOreChecksList(41) = New Setting("DenseVeldspar", CStr(SentSettings.DenseVeldspar))
            IndustryBeltOreChecksList(42) = New Setting("IridescentGneiss", CStr(SentSettings.IridescentGneiss))
            IndustryBeltOreChecksList(43) = New Setting("PrismaticGneiss", CStr(SentSettings.PrismaticGneiss))
            IndustryBeltOreChecksList(44) = New Setting("SilveryOmber", CStr(SentSettings.SilveryOmber))
            IndustryBeltOreChecksList(45) = New Setting("GoldenOmber", CStr(SentSettings.GoldenOmber))
            IndustryBeltOreChecksList(46) = New Setting("MagmaMercoxit", CStr(SentSettings.MagmaMercoxit))
            IndustryBeltOreChecksList(47) = New Setting("VitreousMercoxit", CStr(SentSettings.VitreousMercoxit))

            Call WriteSettingsToFile(IndustryBeltOreChecksFileName, IndustryBeltOreChecksList, "IndustryBeltOreChecksSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving Industry Flip Belt Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetIndustryBeltOreChecksSettings(Belt As BeltType) As IndustryBeltOreChecks
        Select Case Belt
            Case BeltType.Small
                Return IndustryBeltOreChecksSettings1
            Case BeltType.Moderate
                Return IndustryBeltOreChecksSettings2
            Case BeltType.Large
                Return IndustryBeltOreChecksSettings3
            Case BeltType.ExtraLarge
                Return IndustryBeltOreChecksSettings4
            Case BeltType.Giant
                Return IndustryBeltOreChecksSettings5
        End Select
    End Function

#End Region

#Region "Asset Window Settings"

    ' Loads the tab settings
    Public Function LoadAssetWindowSettings(Location As AssetWindow) As AssetWindowSettings
        Dim TempSettings As AssetWindowSettings = Nothing
        Dim AssetWindowFileName As String = ""

        Select Case Location
            Case AssetWindow.ProgramDefault
                AssetWindowFileName = AssetWindowFileNameDefault
            Case AssetWindow.ShoppingList
                AssetWindowFileName = AssetWindowFileNameShoppingList
        End Select

        Try
            If File.Exists(SettingsFolder & AssetWindowFileName) Then

                'Get the settings
                With TempSettings
                    ' Main window
                    .AssetType = CStr(GetSettingValue(AssetWindowFileName, SettingTypes.TypeString, "AssetWindowSettings", "AssetType", DefaultAssetType))
                    .SortbyName = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "SortbyName", DefaultAssetSortbyName))

                    ' Search Settings
                    .ItemFilterText = CStr(GetSettingValue(AssetWindowFileName, SettingTypes.TypeString, "AssetWindowSettings", "ItemFilterText", DefaultAssetItemTextFilter))
                    .AllItems = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "AllItems", DefaultAllItems))
                    .AllRawMats = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "AllRawMats", DefaultAssetItemChecks))
                    .Minerals = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Minerals", DefaultAssetItemChecks))
                    .IceProducts = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "IceProducts", DefaultAssetItemChecks))
                    .Gas = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Gas", DefaultAssetItemChecks))
                    .Misc = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Misc", DefaultAssetItemChecks))
                    .BPCs = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "BPCs", DefaultAssetItemChecks))
                    .AncientRelics = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "AncientRelics", DefaultAssetItemChecks))
                    .AncientSalvage = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "AncientSalvage", DefaultAssetItemChecks))
                    .Salvage = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Salvage", DefaultAssetItemChecks))
                    .StationComponents = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "StationComponents", DefaultAssetItemChecks))
                    .Planetary = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Planetary", DefaultAssetItemChecks))
                    .Datacores = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Datacores", DefaultAssetItemChecks))
                    .Decryptors = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Decryptors", DefaultAssetItemChecks))
                    .RawMats = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "RawMats", DefaultAssetItemChecks))
                    .ProcessedMats = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "ProcessedMats", DefaultAssetItemChecks))
                    .AdvancedMats = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "AdvancedMats", DefaultAssetItemChecks))
                    .MatsandCompounds = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "MatsandCompounds", DefaultAssetItemChecks))
                    .DroneComponents = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "DroneComponents", DefaultAssetItemChecks))
                    .BoosterMats = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "BoosterMats", DefaultAssetItemChecks))
                    .Polymers = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Polymers", DefaultAssetItemChecks))
                    .Asteroids = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Asteroids", DefaultAssetItemChecks))
                    .AllManufacturedItems = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "AllManufacturedItems", DefaultAssetItemChecks))
                    .Ships = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Ships", DefaultAssetItemChecks))
                    .Modules = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Modules", DefaultAssetItemChecks))
                    .Drones = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Drones", DefaultAssetItemChecks))
                    .Boosters = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Boosters", DefaultAssetItemChecks))
                    .Rigs = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Rigs", DefaultAssetItemChecks))
                    .Charges = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Charges", DefaultAssetItemChecks))
                    .Subsystems = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Subsystems", DefaultAssetItemChecks))
                    .Structures = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Structures", DefaultAssetItemChecks))
                    .Tools = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Tools", DefaultAssetItemChecks))
                    .DataInterfaces = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "DataInterfaces", DefaultAssetItemChecks))
                    .CapT2Components = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "CapT2Components", DefaultAssetItemChecks))
                    .CapitalComponents = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "CapitalComponents", DefaultAssetItemChecks))
                    .Components = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Components", DefaultAssetItemChecks))
                    .Hybrid = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Hybrid", DefaultAssetItemChecks))
                    .FuelBlocks = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "FuelBlocks", DefaultAssetItemChecks))
                    .T1 = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "T1", DefaultAssetItemChecks))
                    .T2 = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "T2", DefaultAssetItemChecks))
                    .T3 = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "T3", DefaultAssetItemChecks))
                    .Faction = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Faction", DefaultAssetItemChecks))
                    .Pirate = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Pirate", DefaultAssetItemChecks))
                    .Storyline = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Storyline", DefaultAssetItemChecks))

                    .Celestials = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Celestials", DefaultAssetItemChecks))
                    .Deployables = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Deployables", DefaultAssetItemChecks))
                    .Implants = CBool(GetSettingValue(AssetWindowFileName, SettingTypes.TypeBoolean, "AssetWindowSettings", "Implants", DefaultAssetItemChecks))

                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultAssetWindowSettings(Location)

            End If

        Catch ex As Exception
            MsgBox("An error occured when loading Asset Window Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults                            
            TempSettings = SetDefaultAssetWindowSettings(Location)
        End Try

        ' Save them locally and then export
        Select Case Location
            Case AssetWindow.ProgramDefault
                AssetWindowSettingsDefault = TempSettings
            Case AssetWindow.ShoppingList
                AssetWindowSettingsShoppingList = TempSettings
        End Select

        Return TempSettings

    End Function

    ' Saves the tab settings to XML
    Public Sub SaveAssetWindowSettings(ItemsSelected As AssetWindowSettings, Location As AssetWindow)
        Dim AssetWindowSettingsList(49) As Setting
        Dim AssetWindowFileName As String = ""

        Select Case Location
            Case AssetWindow.ProgramDefault
                AssetWindowFileName = AssetWindowFileNameDefault
            Case AssetWindow.ShoppingList
                AssetWindowFileName = AssetWindowFileNameShoppingList
        End Select

        Try
            AssetWindowSettingsList(0) = New Setting("AllRawMats", CStr(ItemsSelected.AllRawMats))
            AssetWindowSettingsList(1) = New Setting("Minerals", CStr(ItemsSelected.Minerals))
            AssetWindowSettingsList(2) = New Setting("IceProducts", CStr(ItemsSelected.IceProducts))
            AssetWindowSettingsList(3) = New Setting("Gas", CStr(ItemsSelected.Gas))
            AssetWindowSettingsList(4) = New Setting("AncientRelics", CStr(ItemsSelected.AncientRelics))
            AssetWindowSettingsList(5) = New Setting("AncientSalvage", CStr(ItemsSelected.AncientSalvage))
            AssetWindowSettingsList(6) = New Setting("Salvage", CStr(ItemsSelected.Salvage))
            AssetWindowSettingsList(7) = New Setting("StationComponents", CStr(ItemsSelected.StationComponents))
            AssetWindowSettingsList(8) = New Setting("Planetary", CStr(ItemsSelected.Planetary))
            AssetWindowSettingsList(9) = New Setting("Datacores", CStr(ItemsSelected.Datacores))
            AssetWindowSettingsList(10) = New Setting("Decryptors", CStr(ItemsSelected.Decryptors))
            AssetWindowSettingsList(11) = New Setting("RawMats", CStr(ItemsSelected.RawMats))
            AssetWindowSettingsList(12) = New Setting("ProcessedMats", CStr(ItemsSelected.ProcessedMats))
            AssetWindowSettingsList(13) = New Setting("AdvancedMats", CStr(ItemsSelected.AdvancedMats))
            AssetWindowSettingsList(14) = New Setting("MatsandCompounds", CStr(ItemsSelected.MatsandCompounds))
            AssetWindowSettingsList(15) = New Setting("DroneComponents", CStr(ItemsSelected.DroneComponents))
            AssetWindowSettingsList(16) = New Setting("BoosterMats", CStr(ItemsSelected.BoosterMats))
            AssetWindowSettingsList(17) = New Setting("Polymers", CStr(ItemsSelected.Polymers))
            AssetWindowSettingsList(18) = New Setting("AllManufacturedItems", CStr(ItemsSelected.AllManufacturedItems))
            AssetWindowSettingsList(19) = New Setting("Ships", CStr(ItemsSelected.Ships))
            AssetWindowSettingsList(20) = New Setting("Modules", CStr(ItemsSelected.Modules))
            AssetWindowSettingsList(21) = New Setting("Drones", CStr(ItemsSelected.Drones))
            AssetWindowSettingsList(22) = New Setting("Boosters", CStr(ItemsSelected.Boosters))
            AssetWindowSettingsList(23) = New Setting("Rigs", CStr(ItemsSelected.Rigs))
            AssetWindowSettingsList(24) = New Setting("Charges", CStr(ItemsSelected.Charges))
            AssetWindowSettingsList(25) = New Setting("Subsystems", CStr(ItemsSelected.Subsystems))
            AssetWindowSettingsList(26) = New Setting("Structures", CStr(ItemsSelected.Structures))
            AssetWindowSettingsList(27) = New Setting("Tools", CStr(ItemsSelected.Tools))
            AssetWindowSettingsList(28) = New Setting("DataInterfaces", CStr(ItemsSelected.DataInterfaces))
            AssetWindowSettingsList(29) = New Setting("CapT2Components", CStr(ItemsSelected.CapT2Components))
            AssetWindowSettingsList(30) = New Setting("CapitalComponents", CStr(ItemsSelected.CapitalComponents))
            AssetWindowSettingsList(31) = New Setting("Components", CStr(ItemsSelected.Components))
            AssetWindowSettingsList(32) = New Setting("Hybrid", CStr(ItemsSelected.Hybrid))
            AssetWindowSettingsList(33) = New Setting("FuelBlocks", CStr(ItemsSelected.FuelBlocks))
            AssetWindowSettingsList(34) = New Setting("T1", CStr(ItemsSelected.T1))
            AssetWindowSettingsList(35) = New Setting("T2", CStr(ItemsSelected.T2))
            AssetWindowSettingsList(36) = New Setting("T3", CStr(ItemsSelected.T3))
            AssetWindowSettingsList(37) = New Setting("Faction", CStr(ItemsSelected.Faction))
            AssetWindowSettingsList(38) = New Setting("Pirate", CStr(ItemsSelected.Pirate))
            AssetWindowSettingsList(39) = New Setting("Storyline", CStr(ItemsSelected.Storyline))
            AssetWindowSettingsList(40) = New Setting("Asteroids", CStr(ItemsSelected.Asteroids))
            AssetWindowSettingsList(41) = New Setting("Misc", CStr(ItemsSelected.Misc))
            AssetWindowSettingsList(42) = New Setting("ItemFilterText", CStr(ItemsSelected.ItemFilterText))
            AssetWindowSettingsList(43) = New Setting("AllItems", CStr(ItemsSelected.AllItems))

            ' Main window
            AssetWindowSettingsList(44) = New Setting("AssetType", CStr(ItemsSelected.AssetType))
            AssetWindowSettingsList(45) = New Setting("SortbyName", CStr(ItemsSelected.SortbyName))

            AssetWindowSettingsList(46) = New Setting("Celestials", CStr(ItemsSelected.Celestials))
            AssetWindowSettingsList(47) = New Setting("Deployables", CStr(ItemsSelected.Deployables))
            AssetWindowSettingsList(48) = New Setting("Implants", CStr(ItemsSelected.Implants))
            AssetWindowSettingsList(49) = New Setting("BPCs", CStr(ItemsSelected.BPCs))

            Select Case Location
                Case AssetWindow.ProgramDefault
                    Call WriteSettingsToFile(AssetWindowFileName, AssetWindowSettingsList, "DefaultAssetWindowSettings")
                Case AssetWindow.ShoppingList
                    Call WriteSettingsToFile(AssetWindowFileName, AssetWindowSettingsList, "ShoppingListAssetWindowSettings")
            End Select

        Catch ex As Exception
            MsgBox("An error occured when saving Asset Window Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the tab settings
    Public Function GetAssetWindowSettings(Location As AssetWindow) As AssetWindowSettings

        Select Case Location
            Case AssetWindow.ProgramDefault
                Return AssetWindowSettingsDefault
            Case AssetWindow.ShoppingList
                Return AssetWindowSettingsShoppingList
            Case Else
                Return Nothing
        End Select

    End Function

    Public Function SetDefaultAssetWindowSettings(Location As AssetWindow) As AssetWindowSettings
        Dim LocalSettings As AssetWindowSettings = Nothing

        With LocalSettings
            .AssetType = DefaultAssetType
            .SortbyName = DefaultAssetSortbyName

            .ItemFilterText = DefaultAssetItemTextFilter
            .AllItems = DefaultAllItems
            .AllRawMats = DefaultAssetItemChecks
            .Minerals = DefaultAssetItemChecks
            .IceProducts = DefaultAssetItemChecks
            .Gas = DefaultAssetItemChecks
            .Misc = DefaultAssetItemChecks
            .BPCs = DefaultAssetItemChecks
            .AncientRelics = DefaultAssetItemChecks
            .AncientSalvage = DefaultAssetItemChecks
            .Salvage = DefaultAssetItemChecks
            .StationComponents = DefaultAssetItemChecks
            .Planetary = DefaultAssetItemChecks
            .Datacores = DefaultAssetItemChecks
            .Decryptors = DefaultAssetItemChecks
            .RawMats = DefaultAssetItemChecks
            .ProcessedMats = DefaultAssetItemChecks
            .AdvancedMats = DefaultAssetItemChecks
            .MatsandCompounds = DefaultAssetItemChecks
            .DroneComponents = DefaultAssetItemChecks
            .BoosterMats = DefaultAssetItemChecks
            .Polymers = DefaultAssetItemChecks
            .Asteroids = DefaultAssetItemChecks
            .AllManufacturedItems = DefaultAssetItemChecks
            .Ships = DefaultAssetItemChecks
            .Modules = DefaultAssetItemChecks
            .Drones = DefaultAssetItemChecks
            .Boosters = DefaultAssetItemChecks
            .Rigs = DefaultAssetItemChecks
            .Charges = DefaultAssetItemChecks
            .Subsystems = DefaultAssetItemChecks
            .Structures = DefaultAssetItemChecks
            .Tools = DefaultAssetItemChecks
            .DataInterfaces = DefaultAssetItemChecks
            .CapT2Components = DefaultAssetItemChecks
            .CapitalComponents = DefaultAssetItemChecks
            .Components = DefaultAssetItemChecks
            .Hybrid = DefaultAssetItemChecks
            .FuelBlocks = DefaultAssetItemChecks
            .T1 = DefaultAssetItemChecks
            .T2 = DefaultAssetItemChecks
            .T3 = DefaultAssetItemChecks
            .Faction = DefaultAssetItemChecks
            .Pirate = DefaultAssetItemChecks
            .Storyline = DefaultAssetItemChecks
            .Celestials = DefaultAssetItemChecks
            .Deployables = DefaultAssetItemChecks
            .Implants = DefaultAssetItemChecks
        End With

        ' Save locally - Will have more than one
        Select Case Location
            Case AssetWindow.ProgramDefault
                AssetWindowSettingsDefault = LocalSettings
            Case AssetWindow.ShoppingList
                AssetWindowSettingsShoppingList = LocalSettings
        End Select

        Return LocalSettings

    End Function

#End Region

#Region "POS Settings"

    ' Loads the POS tower settings from XML setting file
    Public Function LoadPOSSettings() As PlayerOwnedStationSettings
        Dim TempSettings As PlayerOwnedStationSettings = Nothing

        Try
            If File.Exists(SettingsFolder & POSSettingsFileName) Then
                'Get the settings
                With TempSettings
                    .TowerRaceID = CInt(GetSettingValue(POSSettingsFileName, SettingTypes.TypeInteger, "POSSettings", "TowerRaceID", DefaultTowerRaceID))
                    .TowerName = CStr(GetSettingValue(POSSettingsFileName, SettingTypes.TypeString, "POSSettings", "TowerName", DefaultTowerName))
                    .CostperHour = CDbl(GetSettingValue(POSSettingsFileName, SettingTypes.TypeDouble, "POSSettings", "CostperHour", DefaultCostperHour))
                    .MECostperSecond = CDbl(GetSettingValue(POSSettingsFileName, SettingTypes.TypeDouble, "POSSettings", "MECostperSecond", DefaultMECostperSecond))
                    .TECostperSecond = CDbl(GetSettingValue(POSSettingsFileName, SettingTypes.TypeDouble, "POSSettings", "TECostperSecond", DefaultTECostperSecond))
                    .InventionCostperSecond = CDbl(GetSettingValue(POSSettingsFileName, SettingTypes.TypeDouble, "POSSettings", "InventionCostperSecond", DefaultInventionCostperSecond))
                    .CopyCostperSecond = CDbl(GetSettingValue(POSSettingsFileName, SettingTypes.TypeDouble, "POSSettings", "CopyCostperSecond", DefaultCopyCostperSecond))
                    .TowerType = CStr(GetSettingValue(POSSettingsFileName, SettingTypes.TypeString, "POSSettings", "TowerType", DefaultTowerType))
                    .TowerSize = CStr(GetSettingValue(POSSettingsFileName, SettingTypes.TypeString, "POSSettings", "TowerSize", DefaultTowerSize))
                    .FuelBlockBuild = CBool(GetSettingValue(POSSettingsFileName, SettingTypes.TypeBoolean, "POSSettings", "FuelBlockBuild", DefaultFuelBlockBuild))
                    .NumAdvLabs = CInt(GetSettingValue(POSSettingsFileName, SettingTypes.TypeInteger, "POSSettings", "NumAdvLabs", DefaultNumAdvLabs))
                    .NumMobileLabs = CInt(GetSettingValue(POSSettingsFileName, SettingTypes.TypeInteger, "POSSettings", "NumMobileLabs", DefaultNumMobileLabs))
                    .NumHyasyodaLabs = CInt(GetSettingValue(POSSettingsFileName, SettingTypes.TypeInteger, "POSSettings", "NumHyasyodaLabs", DefaultNumHyasyodaLabs))
                    .CharterCost = CDbl(GetSettingValue(POSSettingsFileName, SettingTypes.TypeInteger, "POSSettings", "CharterCost", DefaultCharterCost))
                End With

            Else
                ' Load defaults 
                TempSettings = SetDefaultPOSSettings()
            End If
        Catch ex As Exception
            MsgBox("An error occured when loading POS Tower Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultPOSSettings()
        End Try

        ' Save them locally and then export
        POSSettings = TempSettings

        Return TempSettings

    End Function

    ' Load defaults 
    Public Function SetDefaultPOSSettings() As PlayerOwnedStationSettings
        Dim TempSettings As PlayerOwnedStationSettings = Nothing

        ' Load defaults 
        TempSettings.TowerRaceID = DefaultTowerRaceID
        TempSettings.TowerName = DefaultTowerName
        TempSettings.CostperHour = DefaultCostperHour
        TempSettings.MECostperSecond = DefaultMECostperSecond
        TempSettings.TECostperSecond = DefaultTECostperSecond
        TempSettings.InventionCostperSecond = DefaultInventionCostperSecond
        TempSettings.CopyCostperSecond = DefaultCopyCostperSecond
        TempSettings.TowerType = DefaultTowerType
        TempSettings.TowerSize = DefaultTowerSize
        TempSettings.FuelBlockBuild = DefaultFuelBlockBuild
        TempSettings.NumAdvLabs = DefaultNumAdvLabs
        TempSettings.NumMobileLabs = DefaultNumMobileLabs
        TempSettings.NumHyasyodaLabs = DefaultNumHyasyodaLabs
        TempSettings.CharterCost = DefaultCharterCost

        POSSettings = TempSettings
        Return TempSettings

    End Function

    ' Saves the POS Settings to XML
    Public Sub SavePOSSettings(SentSettings As PlayerOwnedStationSettings)
        Dim POSSettingsList(13) As Setting

        Try
            POSSettingsList(0) = New Setting("TowerRaceID", CStr(SentSettings.TowerRaceID))
            POSSettingsList(1) = New Setting("TowerName", CStr(SentSettings.TowerName))
            POSSettingsList(2) = New Setting("CostperHour", CStr(SentSettings.CostperHour))
            POSSettingsList(3) = New Setting("MECostperSecond", CStr(SentSettings.MECostperSecond))
            POSSettingsList(4) = New Setting("TECostperSecond", CStr(SentSettings.TECostperSecond))
            POSSettingsList(5) = New Setting("InventionCostperSecond", CStr(SentSettings.InventionCostperSecond))
            POSSettingsList(6) = New Setting("CopyCostperSecond", CStr(SentSettings.CopyCostperSecond))
            POSSettingsList(7) = New Setting("TowerType", CStr(SentSettings.TowerType))
            POSSettingsList(8) = New Setting("TowerSize", CStr(SentSettings.TowerSize))
            POSSettingsList(9) = New Setting("FuelBlockBuild", CStr(SentSettings.FuelBlockBuild))
            POSSettingsList(10) = New Setting("NumAdvLabs", CStr(SentSettings.NumAdvLabs))
            POSSettingsList(11) = New Setting("NumMobileLabs", CStr(SentSettings.NumMobileLabs))
            POSSettingsList(12) = New Setting("NumHyasyodaLabs", CStr(SentSettings.NumHyasyodaLabs))
            POSSettingsList(13) = New Setting("CharterCost", CStr(SentSettings.CharterCost))

            Call WriteSettingsToFile(POSSettingsFileName, POSSettingsList, "POSSettings")

        Catch ex As Exception
            MsgBox("An error occured when saving POS Tower Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the POS Tower Settings
    Public Function GetPOSSettings() As PlayerOwnedStationSettings
        Return POSSettings
    End Function

#End Region

#Region "Team Settings"

    ' Loads the Team settings from XML setting file
    Public Function LoadTeamSettings(IndustryTeamType As TeamType, Tab As String) As TeamSettings
        Dim TempSettings As TeamSettings = Nothing
        Dim TeamFileName As String = ""

        Select Case IndustryTeamType
            Case TeamType.Manufacturing
                TeamFileName = Tab & ManufacturingTeamSettingsFileName
            Case TeamType.ComponentManufacturing
                TeamFileName = Tab & ComponentManufacturingTeamSettingsFileName
            Case TeamType.Copy
                TeamFileName = Tab & CopyTeamSettingsFileName
            Case TeamType.Invention
                TeamFileName = Tab & InventionTeamSettingsFileName
        End Select

        Try
            If File.Exists(SettingsFolder & TeamFileName) Then
                'Get the settings
                With TempSettings
                    Select Case IndustryTeamType
                        Case TeamType.Manufacturing
                            .TeamID = CInt(GetSettingValue(TeamFileName, SettingTypes.TypeInteger, Tab & "ManufacturingTeamSettings", "TeamID", DefaultTeamID))
                        Case TeamType.ComponentManufacturing
                            .TeamID = CInt(GetSettingValue(TeamFileName, SettingTypes.TypeInteger, Tab & "ComponentManufacturingTeamSettings", "TeamID", DefaultTeamID))
                        Case TeamType.Copy
                            .TeamID = CInt(GetSettingValue(TeamFileName, SettingTypes.TypeInteger, Tab & "CopyTeamSettings", "TeamID", DefaultTeamID))
                        Case TeamType.Invention
                            .TeamID = CInt(GetSettingValue(TeamFileName, SettingTypes.TypeInteger, Tab & "InventionTeamSettings", "TeamID", DefaultTeamID))
                    End Select
                End With
            Else
                ' Load defaults 
                TempSettings = SetDefaultTeamSettings(IndustryTeamType, Tab)
            End If
        Catch ex As Exception
            MsgBox("An error occured when loading Team Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetDefaultTeamSettings(IndustryTeamType, Tab)
        End Try

        TempSettings.TeamTab = Tab

        ' Save them locally and then export
        Select Case IndustryTeamType
            Case TeamType.Manufacturing
                BPManufacturingTeamSettings = TempSettings
            Case TeamType.ComponentManufacturing
                BPComponentManufacturingTeamSettings = TempSettings
            Case TeamType.Copy
                BPCopyTeamSettings = TempSettings
            Case TeamType.Invention
                BPInventionTeamSettings = TempSettings
        End Select

        Return TempSettings

    End Function

    ' Saves the Team Settings to XML
    Public Sub SaveTeamSettings(SentSettings As TeamSettings, IndustryTeamType As TeamType)
        Dim TeamSettingsList(0) As Setting

        Try
            TeamSettingsList(0) = New Setting("TeamID", CStr(SentSettings.TeamID))

            Select Case IndustryTeamType
                Case TeamType.Manufacturing
                    Call WriteSettingsToFile(SentSettings.TeamTab & ManufacturingTeamSettingsFileName, TeamSettingsList, SentSettings.TeamTab & "ManufacturingTeamSettings")
                Case TeamType.ComponentManufacturing
                    Call WriteSettingsToFile(SentSettings.TeamTab & ComponentManufacturingTeamSettingsFileName, TeamSettingsList, SentSettings.TeamTab & "ComponentManufacturingTeamSettings")
                Case TeamType.Copy
                    Call WriteSettingsToFile(SentSettings.TeamTab & CopyTeamSettingsFileName, TeamSettingsList, SentSettings.TeamTab & "CopyTeamSettings")
                Case TeamType.Invention
                    Call WriteSettingsToFile(SentSettings.TeamTab & InventionTeamSettingsFileName, TeamSettingsList, SentSettings.TeamTab & "InventionTeamSettings")
            End Select

        Catch ex As Exception
            MsgBox("An error occured when saving Team Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the Team Settings
    Public Function GetTeamSettings(IndustryTeamType As TeamType, Tab As String) As TeamSettings

        If Tab = BPTab Then
            Select Case IndustryTeamType
                Case TeamType.Manufacturing
                    Return BPManufacturingTeamSettings
                Case TeamType.ComponentManufacturing
                    Return BPComponentManufacturingTeamSettings
                Case TeamType.Copy
                    Return BPCopyTeamSettings
                Case TeamType.Invention
                    Return BPInventionTeamSettings
            End Select
        Else
            Select Case IndustryTeamType
                Case TeamType.Manufacturing
                    Return CalcManufacturingTeamSettings
                Case TeamType.ComponentManufacturing
                    Return CalcComponentManufacturingTeamSettings
                Case TeamType.Copy
                    Return CalcCopyTeamSettings
                Case TeamType.Invention
                    Return CalcInventionTeamSettings
            End Select
        End If

        Return Nothing

    End Function

    ' Load defaults 
    Public Function SetDefaultTeamSettings(IndustryTeamType As TeamType, Tab As String) As TeamSettings
        Dim TempSettings As TeamSettings = Nothing

        ' Load defaults 
        TempSettings.TeamID = DefaultTeamID

        If Tab = BPTab Then
            Select Case IndustryTeamType
                Case TeamType.Manufacturing
                    BPManufacturingTeamSettings = TempSettings
                Case TeamType.ComponentManufacturing
                    BPComponentManufacturingTeamSettings = TempSettings
                Case TeamType.Copy
                    BPCopyTeamSettings = TempSettings
                Case TeamType.Invention
                    BPInventionTeamSettings = TempSettings
            End Select
        Else
            Select Case IndustryTeamType
                Case TeamType.Manufacturing
                    CalcManufacturingTeamSettings = TempSettings
                Case TeamType.ComponentManufacturing
                    CalcComponentManufacturingTeamSettings = TempSettings
                Case TeamType.Copy
                    CalcCopyTeamSettings = TempSettings
                Case TeamType.Invention
                    CalcInventionTeamSettings = TempSettings
            End Select
        End If

        TempSettings.TeamTab = Tab

        Return TempSettings

    End Function

#End Region

#Region "Facility Settings"

    ' Loads the Facility settings from XML setting file
    Public Function LoadFacilitySettings(ProductionType As IndustryType, Tab As String) As FacilitySettings
        Dim TempSettings As FacilitySettings = Nothing
        Dim FacilityFileName As String = ""

        Select Case ProductionType
            Case IndustryType.Manufacturing
                FacilityFileName = ManufacturingFacilitySettingsFileName
            Case IndustryType.ComponentManufacturing
                FacilityFileName = ComponentsManufacturingFacilitySettingsFileName
            Case IndustryType.CapitalComponentManufacturing
                FacilityFileName = CapitalComponentsManufacturingFacilitySettingsFileName
            Case IndustryType.CapitalManufacturing
                FacilityFileName = CapitalManufacturingFacilitySettingsFileName
            Case IndustryType.SuperManufacturing
                FacilityFileName = SuperCapitalManufacturingFacilitySettingsFileName
            Case IndustryType.BoosterManufacturing
                FacilityFileName = BoosterManufacturingFacilitySettingsFileName
            Case IndustryType.T3CruiserManufacturing
                FacilityFileName = T3CruiserManufacturingFacilitySettingsFileName
            Case IndustryType.T3DestroyerManufacturing
                FacilityFileName = T3DestroyerManufacturingFacilitySettingsFileName
            Case IndustryType.SubsystemManufacturing
                FacilityFileName = SubsystemManufacturingFacilitySettingsFileName
            Case IndustryType.Copying
                FacilityFileName = CopyFacilitySettingsFileName
            Case IndustryType.Invention
                FacilityFileName = InventionFacilitySettingsFileName
            Case IndustryType.T3Invention
                FacilityFileName = T3InventionFacilitySettingsFileName
            Case IndustryType.NoPOSManufacturing
                FacilityFileName = NoPoSFacilitySettingsFileName
            Case IndustryType.POSFuelBlockManufacturing
                FacilityFileName = POSFuelBlockFacilitySettingsFileName
            Case IndustryType.POSLargeShipManufacturing
                FacilityFileName = POSLargeShipFacilitySettingsFileName
            Case IndustryType.POSModuleManufacturing
                FacilityFileName = POSModuleFacilitySettingsFileName
        End Select

        FacilityFileName = Tab & FacilityFileName

        Try
            If File.Exists(SettingsFolder & FacilityFileName) Then
                'Get the settings
                With TempSettings
                    Select Case ProductionType
                        Case IndustryType.Manufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "ManufacturingFacilitySettings", "Facility", DefaultBPManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "ManufacturingFacilitySettings", "FacilityType", DefaultManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "ManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "ManufacturingFacilitySettings", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "ManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "ManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "ManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "ManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "ManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "ManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "ManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "ManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "ManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "ManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.ComponentManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "ComponentsManufacturingFacilitySettings", "Facility", DefaultComponentManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "ComponentsManufacturingFacilitySettings", "FacilityType", DefaultComponentManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "ComponentsManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "ComponentsManufacturingFacilitySettings", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "ComponentsManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "ComponentsManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "ComponentsManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "ComponentsManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "ComponentsManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "ComponentsManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "ComponentsManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "ComponentsManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "ComponentsManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "ComponentsManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.CapitalComponentManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CapitalComponentsManufacturingFacilitySettings", "Facility", DefaultCapitalComponentManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CapitalComponentsManufacturingFacilitySettings", "FacilityType", DefaultCapitalComponentManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "CapitalComponentsManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "CapitalComponentsManufacturingFacilitySettings", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CapitalComponentsManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CapitalComponentsManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "CapitalComponentsManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CapitalComponentsManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "CapitalComponentsManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CapitalComponentsManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CapitalComponentsManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CapitalComponentsManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CapitalComponentsManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CapitalComponentsManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.CapitalManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CapitalManufacturingFacilitySettings", "Facility", DefaultCapitalManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CapitalManufacturingFacilitySettings", "FacilityType", DefaultCapitalManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "CapitalManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "CapitalManufacturingFacilitySettings", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CapitalManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CapitalManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "CapitalManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CapitalManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "CapitalManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CapitalManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CapitalManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CapitalManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CapitalManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CapitalManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.SuperManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "SuperManufacturingFacilitySettings", "Facility", DefaultSuperManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "SuperManufacturingRAM_INSTALLATION_TYPE_CONTENTS", "FacilityType", DefaultSuperManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "SuperManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "SuperManufacturingFacilitySettings", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "SuperManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "SuperManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "SuperManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "SuperManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "SuperManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "SuperManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "SuperManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "SuperManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "SuperManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "SuperManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.BoosterManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "BoosterManufacturingFacilitySettings", "Facility", DefaultBoosterManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "BoosterManufacturingFacilitySettings", "FacilityType", DefaultBoosterManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "BoosterManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "BoosterManufacturingFacilitySettings", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "BoosterManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "BoosterManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "BoosterManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "BoosterManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "BoosterManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "BoosterManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "BoosterManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "BoosterManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "BoosterManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "BoosterManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.T3CruiserManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3CruiserManufacturingFacilitySettings", "Facility", DefaultT3CruiserManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3CruiserManufacturingFacilitySettings", "FacilityType", DefaultT3CruiserManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "T3CruiserManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "T3CruiserManufacturingFacilitySettings", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3CruiserManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3CruiserManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "T3CruiserManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3CruiserManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "T3CruiserManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3CruiserManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3CruiserManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3CruiserManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3CruiserManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3CruiserManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.T3DestroyerManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3DestroyerManufacturingFacilitySettings", "Facility", DefaultT3DestroyerManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3DestroyerManufacturingFacilitySettings", "FacilityType", DefaultT3DestroyerManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "T3DestroyerManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "T3DestroyerManufacturingFacilitySettings", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3DestroyerManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3DestroyerManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "T3DestroyerManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3DestroyerManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "T3DestroyerManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3DestroyerManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3DestroyerManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3DestroyerManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3DestroyerManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3DestroyerManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.SubsystemManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "SubsystemManufacturingFacilitySettings", "Facility", DefaultSubsystemManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "SubsystemManufacturingFacilitySettings", "FacilityType", DefaultSubsystemManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "SubsystemManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "SubsystemManufacturing", "ProductionType", IndustryType.Manufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "SubsystemManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "SubsystemManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "SubsystemManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "SubsystemManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "SubsystemManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "SubsystemManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "SubsystemManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "SubsystemManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "SubsystemManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "SubsystemManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.Copying
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CopyFacilitySettings", "Facility", DefaultCopyFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CopyFacilitySettings", "FacilityType", DefaultCopyFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "CopyFacilitySettings", "ActivityID", IndustryActivities.Copying))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "CopyFacilitySettings", "ProductionType", IndustryType.Copying), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CopyFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CopyFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "CopyFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CopyFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "CopyFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "CopyFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "CopyFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CopyFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CopyFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "CopyFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.Invention
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "InventionFacilitySettings", "Facility", DefaultInventionFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "InventionFacilitySettings", "FacilityType", DefaultInventionFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "InventionFacilitySettings", "ActivityID", IndustryActivities.Invention))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "InventionFacilitySettings", "ProductionType", IndustryType.Invention), IndustryType)
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "InventionFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "InventionFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "InventionFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "InventionFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "InventionFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "InventionFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "InventionFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "InventionFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "InventionFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "InventionFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.T3Invention
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3InventionFacilitySettings", "Facility", DefaultT3InventionFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3InventionFacilitySettings", "FacilityType", DefaultT3InventionFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "T3InventionFacilitySettings", "ActivityID", IndustryActivities.Invention))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "T3InventionFacilitySettings", "ProductionType", IndustryType.Invention), IndustryType)
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3InventionFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3InventionFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "T3InventionFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3InventionFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "T3InventionFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "T3InventionFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "T3InventionFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3InventionFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3InventionFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "T3InventionFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.NoPOSManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "NoPOSFacilitySettings", "Facility", DefaultManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "NoPOSFacilitySettings", "FacilityType", DefaultManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "NoPOSFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "NoPOSFacilitySettings", "ProductionType", IndustryType.NoPOSManufacturing), IndustryType)
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "NoPOSFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "NoPOSFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "NoPOSFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "NoPOSFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "NoPOSFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "NoPOSFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "NoPOSFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "NoPOSFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "NoPOSFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "NoPOSFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.POSFuelBlockManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSFuelBlockManufacturingFacilitySettings", "Facility", DefaultPOSFuelBlockManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSFuelBlockManufacturingFacilitySettings", "FacilityType", DefaultPOSFuelBlockManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "POSFuelBlockManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "POSFuelBlockManufacturingFacilitySettings", "ProductionType", IndustryType.POSFuelBlockManufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSFuelBlockManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSFuelBlockManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "POSFuelBlockManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSFuelBlockManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "POSFuelBlockManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSFuelBlockManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSFuelBlockManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSFuelBlockManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSFuelBlockManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSFuelBlockManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.POSModuleManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSModuleManufacturingFacilitySettings", "Facility", DefaultPOSModuleManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSModuleManufacturingFacilitySettings", "FacilityType", DefaultPOSModuleManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "POSModuleManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "POSModuleManufacturingFacilitySettings", "ProductionType", IndustryType.POSModuleManufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSModuleManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSModuleManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "POSModuleManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSModuleManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "POSModuleManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSModuleManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSModuleManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSModuleManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSModuleManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSModuleManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                        Case IndustryType.POSLargeShipManufacturing
                            .Facility = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSLargeShipManufacturingFacilitySettings", "Facility", DefaultPOSLargeShipManufacturingFacility))
                            .FacilityType = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSLargeShipManufacturingFacilitySettings", "FacilityType", DefaultPOSLargeShipManufacturingFacilityType))
                            .ActivityID = CInt(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "POSLargeShipManufacturingFacilitySettings", "ActivityID", IndustryActivities.Manufacturing))
                            .ProductionType = CType(GetSettingValue(FacilityFileName, SettingTypes.TypeInteger, Tab & "POSLargeShipManufacturingFacilitySettings", "ProductionType", IndustryType.POSLargeShipManufacturing), IndustryType)
                            .MaterialMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSLargeShipManufacturingFacilitySettings", "MaterialMultiplier", FacilityDefaultMM))
                            .TimeMultiplier = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSLargeShipManufacturingFacilitySettings", "TimeMultiplier", FacilityDefaultTM))
                            .SolarSystemID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "POSLargeShipManufacturingFacilitySettings", "SolarSystemID", FacilityDefaultSolarSystemID))
                            .SolarSystemName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSLargeShipManufacturingFacilitySettings", "SolarSystemName", FacilityDefaultSolarSystem))
                            .RegionID = CLng(GetSettingValue(FacilityFileName, SettingTypes.TypeLong, Tab & "POSLargeShipManufacturingFacilitySettings", "RegionID", FacilityDefaultRegionID))
                            .RegionName = CStr(GetSettingValue(FacilityFileName, SettingTypes.TypeString, Tab & "POSLargeShipManufacturingFacilitySettings", "RegionName", FacilityDefaultRegion))
                            .ActivityCostperSecond = CDbl(GetSettingValue(FacilityFileName, SettingTypes.TypeDouble, Tab & "POSLargeShipManufacturingFacilitySettings", "ActivityCostperSecond", FacilityDefaultActivityCostperSecond))
                            .IncludeActivityUsage = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSLargeShipManufacturingFacilitySettings", "IncludeActivityUsage", FacilityDefaultIncludeUsage))
                            .IncludeActivityCost = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSLargeShipManufacturingFacilitySettings", "IncludeActivityCost", FacilityDefaultIncludeCost))
                            .IncludeActivityTime = CBool(GetSettingValue(FacilityFileName, SettingTypes.TypeBoolean, Tab & "POSLargeShipManufacturingFacilitySettings", "IncludeActivityTime", FacilityDefaultIncludeTime))
                    End Select
                End With
            Else
                ' Load defaults 
                TempSettings = SetFacilityDefaultSettings(ProductionType)
            End If
        Catch ex As Exception
            MsgBox("An error occured when loading Facility Tower Settings. Error: " & Err.Description & vbCrLf & "Default settings were loaded.", vbExclamation, Application.ProductName)
            ' Load defaults 
            TempSettings = SetFacilityDefaultSettings(ProductionType)
        End Try

        ' Save them locally and then export
        Select Case ProductionType
            Case IndustryType.Manufacturing
                ManufacturingFacilitySettings = TempSettings
            Case IndustryType.ComponentManufacturing
                ComponentsManufacturingFacilitySettings = TempSettings
            Case IndustryType.CapitalComponentManufacturing
                CapitalComponentsManufacturingFacilitySettings = TempSettings
            Case IndustryType.CapitalManufacturing
                CapitalManufacturingFacilitySettings = TempSettings
            Case IndustryType.SuperManufacturing
                SuperManufacturingFacilitySettings = TempSettings
            Case IndustryType.BoosterManufacturing
                BoosterManufacturingFacilitySettings = TempSettings
            Case IndustryType.T3CruiserManufacturing
                T3CruiserManufacturingFacilitySettings = TempSettings
            Case IndustryType.T3DestroyerManufacturing
                T3DestroyerManufacturingFacilitySettings = TempSettings
            Case IndustryType.SubsystemManufacturing
                SubsystemManufacturingFacilitySettings = TempSettings
            Case IndustryType.Copying
                CopyFacilitySettings = TempSettings
            Case IndustryType.Invention
                InventionFacilitySettings = TempSettings
            Case IndustryType.T3Invention
                T3InventionFacilitySettings = TempSettings
            Case IndustryType.NoPOSManufacturing
                NoPoSFacilitySettings = TempSettings
            Case IndustryType.POSFuelBlockManufacturing
                POSFuelBlockFacilitySettings = TempSettings
            Case IndustryType.POSLargeShipManufacturing
                POSLargeShipFacilitySettings = TempSettings
            Case IndustryType.POSModuleManufacturing
                POSModuleFacilitySettings = TempSettings
        End Select

        Return TempSettings

    End Function

    ' Saves the Facility Settings to XML
    Public Sub FacilitySaveSettings(SentSettings As FacilitySettings, ProductionType As IndustryType, Tab As String)
        Dim FacilitySettingsList(13) As Setting

        Try
            FacilitySettingsList(0) = New Setting("Facility", CStr(SentSettings.Facility))
            FacilitySettingsList(1) = New Setting("FacilityType", CStr(SentSettings.FacilityType))
            FacilitySettingsList(2) = New Setting("ActivityID", CStr(SentSettings.ActivityID))
            FacilitySettingsList(3) = New Setting("ProductionType", CStr(SentSettings.ProductionType))
            FacilitySettingsList(4) = New Setting("MaterialMultiplier", CStr(SentSettings.MaterialMultiplier))
            FacilitySettingsList(5) = New Setting("TimeMultiplier", CStr(SentSettings.TimeMultiplier))
            FacilitySettingsList(6) = New Setting("ActivityCostperSecond", CStr(SentSettings.ActivityCostperSecond))
            FacilitySettingsList(7) = New Setting("SolarSystemID", CStr(SentSettings.SolarSystemID))
            FacilitySettingsList(8) = New Setting("SolarSystemName", CStr(SentSettings.SolarSystemName))
            FacilitySettingsList(9) = New Setting("RegionID", CStr(SentSettings.RegionID))
            FacilitySettingsList(10) = New Setting("RegionName", CStr(SentSettings.RegionName))
            FacilitySettingsList(11) = New Setting("IncludeActivityUsage", CStr(SentSettings.IncludeActivityUsage))
            FacilitySettingsList(12) = New Setting("IncludeActivityCost", CStr(SentSettings.IncludeActivityCost))
            FacilitySettingsList(13) = New Setting("IncludeActivityTime", CStr(SentSettings.IncludeActivityTime))

            Select Case ProductionType
                Case IndustryType.Manufacturing
                    Call WriteSettingsToFile(Tab & ManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "ManufacturingFacilitySettings")
                Case IndustryType.ComponentManufacturing
                    Call WriteSettingsToFile(Tab & ComponentsManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "ComponentsManufacturingFacilitySettings")
                Case IndustryType.CapitalComponentManufacturing
                    Call WriteSettingsToFile(Tab & CapitalComponentsManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "CapitalComponentsManufacturingFacilitySettings")
                Case IndustryType.SubsystemManufacturing
                    Call WriteSettingsToFile(Tab & SubsystemManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "SubsystemManufacturingFacilitySettings")
                Case IndustryType.SuperManufacturing
                    Call WriteSettingsToFile(Tab & SuperCapitalManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "SuperManufacturingFacilitySettings")
                Case IndustryType.T3CruiserManufacturing
                    Call WriteSettingsToFile(Tab & T3CruiserManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "T3CruiserManufacturingFacilitySettings")
                Case IndustryType.T3DestroyerManufacturing
                    Call WriteSettingsToFile(Tab & T3DestroyerManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "T3DestroyerManufacturingFacilitySettings")
                Case IndustryType.BoosterManufacturing
                    Call WriteSettingsToFile(Tab & BoosterManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "BoosterManufacturingFacilitySettings")
                Case IndustryType.CapitalManufacturing
                    Call WriteSettingsToFile(Tab & CapitalManufacturingFacilitySettingsFileName, FacilitySettingsList, Tab & "CapitalManufacturingFacilitySettings")
                Case IndustryType.Copying
                    Call WriteSettingsToFile(Tab & CopyFacilitySettingsFileName, FacilitySettingsList, Tab & "CopyFacilitySettings")
                Case IndustryType.Invention
                    Call WriteSettingsToFile(Tab & InventionFacilitySettingsFileName, FacilitySettingsList, Tab & "InventionFacilitySettings")
                Case IndustryType.T3Invention
                    Call WriteSettingsToFile(Tab & T3InventionFacilitySettingsFileName, FacilitySettingsList, Tab & "T3InventionFacilitySettings")
                Case IndustryType.NoPOSManufacturing
                    Call WriteSettingsToFile(Tab & NoPoSFacilitySettingsFileName, FacilitySettingsList, Tab & "NoPOSFacilitySettings")
                Case IndustryType.POSFuelBlockManufacturing
                    Call WriteSettingsToFile(Tab & POSFuelBlockFacilitySettingsFileName, FacilitySettingsList, Tab & "POSFuelBlockManufacturingFacilitySettings")
                Case IndustryType.POSLargeShipManufacturing
                    Call WriteSettingsToFile(Tab & POSLargeShipFacilitySettingsFileName, FacilitySettingsList, Tab & "POSLargeShipManufacturingFacilitySettings")
                Case IndustryType.POSModuleManufacturing
                    Call WriteSettingsToFile(Tab & POSModuleFacilitySettingsFileName, FacilitySettingsList, Tab & "POSModuleManufacturingFacilitySettings")
            End Select

        Catch ex As Exception
            MsgBox("An error occured when saving Manufacturing Facility Settings. Error: " & Err.Description & vbCrLf & "Settings not saved.", vbExclamation, Application.ProductName)
        End Try

    End Sub

    ' Returns the Facility Settings
    Public Function GetFacilitySettings(ProductionType As IndustryType) As FacilitySettings

        Select Case ProductionType
            Case IndustryType.Manufacturing
                Return ManufacturingFacilitySettings
            Case IndustryType.BoosterManufacturing
                Return BoosterManufacturingFacilitySettings
            Case IndustryType.CapitalManufacturing
                Return CapitalManufacturingFacilitySettings
            Case IndustryType.SubsystemManufacturing
                Return SubsystemManufacturingFacilitySettings
            Case IndustryType.SuperManufacturing
                Return SuperManufacturingFacilitySettings
            Case IndustryType.T3CruiserManufacturing
                Return T3CruiserManufacturingFacilitySettings
            Case IndustryType.T3DestroyerManufacturing
                Return T3DestroyerManufacturingFacilitySettings
            Case IndustryType.ComponentManufacturing
                Return ComponentsManufacturingFacilitySettings
            Case IndustryType.CapitalComponentManufacturing
                Return CapitalComponentsManufacturingFacilitySettings
            Case IndustryType.Copying
                Return CopyFacilitySettings
            Case IndustryType.Invention
                Return InventionFacilitySettings
            Case IndustryType.T3Invention
                Return T3InventionFacilitySettings
            Case IndustryType.NoPOSManufacturing
                Return NoPoSFacilitySettings
            Case IndustryType.POSFuelBlockManufacturing
                Return POSFuelBlockFacilitySettings
            Case IndustryType.POSLargeShipManufacturing
                Return POSLargeShipFacilitySettings
            Case IndustryType.POSModuleManufacturing
                Return POSModuleFacilitySettings
        End Select

        Return Nothing

    End Function

    ' Load defaults 
    Public Function SetFacilityDefaultSettings(ProductionType As IndustryType) As FacilitySettings
        Dim TempSettings As FacilitySettings = Nothing

        ' These are all the same regardless
        TempSettings.MaterialMultiplier = FacilityDefaultMM
        TempSettings.TimeMultiplier = FacilityDefaultTM
        TempSettings.ActivityCostperSecond = FacilityDefaultActivityCostperSecond
        TempSettings.IncludeActivityUsage = FacilityDefaultIncludeUsage
        TempSettings.IncludeActivityCost = FacilityDefaultIncludeCost
        TempSettings.IncludeActivityTime = FacilityDefaultIncludeTime

        ' For POS settings, use a null setting for boosters and supers, else use highsec (Jita)
        Select Case ProductionType
            Case IndustryType.SuperManufacturing, IndustryType.BoosterManufacturing
                TempSettings.SolarSystemID = DefaultNullFacilitySolarSystemID
                TempSettings.SolarSystemName = DefaultNullFacilitySolarSystem
                TempSettings.RegionID = DefaultNullFacilityRegionID
                TempSettings.RegionName = DefaultNullFacilityRegion
            Case Else
                TempSettings.SolarSystemID = FacilityDefaultSolarSystemID
                TempSettings.SolarSystemName = FacilityDefaultSolarSystem
                TempSettings.RegionID = FacilityDefaultRegionID
                TempSettings.RegionName = FacilityDefaultRegion
        End Select

        ' Load defaults 
        Select Case ProductionType
            Case IndustryType.Manufacturing
                TempSettings.Facility = DefaultManufacturingFacility
                TempSettings.FacilityType = DefaultManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.Manufacturing
                ManufacturingFacilitySettings = TempSettings
            Case IndustryType.SuperManufacturing
                TempSettings.Facility = DefaultSuperManufacturingFacility
                TempSettings.FacilityType = DefaultSuperManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.SuperManufacturing
                SuperManufacturingFacilitySettings = TempSettings
            Case IndustryType.BoosterManufacturing
                TempSettings.Facility = DefaultBoosterManufacturingFacility
                TempSettings.FacilityType = DefaultBoosterManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.BoosterManufacturing
                BoosterManufacturingFacilitySettings = TempSettings
            Case IndustryType.SubsystemManufacturing
                TempSettings.Facility = DefaultSubsystemManufacturingFacility
                TempSettings.FacilityType = DefaultSubsystemManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.SubsystemManufacturing
                SubsystemManufacturingFacilitySettings = TempSettings
            Case IndustryType.T3CruiserManufacturing
                TempSettings.Facility = DefaultT3CruiserManufacturingFacility
                TempSettings.FacilityType = DefaultT3CruiserManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.T3CruiserManufacturing
                T3CruiserManufacturingFacilitySettings = TempSettings
            Case IndustryType.T3DestroyerManufacturing
                TempSettings.Facility = DefaultT3DestroyerManufacturingFacility
                TempSettings.FacilityType = DefaultT3DestroyerManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.T3DestroyerManufacturing
                T3DestroyerManufacturingFacilitySettings = TempSettings
            Case IndustryType.CapitalManufacturing
                TempSettings.Facility = DefaultCapitalManufacturingFacility
                TempSettings.FacilityType = DefaultCapitalManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.CapitalManufacturing
                CapitalManufacturingFacilitySettings = TempSettings
            Case IndustryType.ComponentManufacturing
                TempSettings.Facility = DefaultComponentManufacturingFacility
                TempSettings.FacilityType = DefaultComponentManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.ComponentManufacturing
                ComponentsManufacturingFacilitySettings = TempSettings
            Case IndustryType.CapitalComponentManufacturing
                TempSettings.Facility = DefaultCapitalComponentManufacturingFacility
                TempSettings.FacilityType = DefaultCapitalComponentManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.CapitalComponentManufacturing
                CapitalComponentsManufacturingFacilitySettings = TempSettings
            Case IndustryType.Copying
                TempSettings.Facility = DefaultCopyFacility
                TempSettings.FacilityType = DefaultCopyFacilityType
                TempSettings.ActivityID = IndustryActivities.Copying
                TempSettings.ProductionType = IndustryType.Copying
                CopyFacilitySettings = TempSettings
            Case IndustryType.Invention
                TempSettings.Facility = DefaultInventionFacility
                TempSettings.FacilityType = DefaultInventionFacilityType
                TempSettings.ActivityID = IndustryActivities.Invention
                TempSettings.ProductionType = IndustryType.Invention
                InventionFacilitySettings = TempSettings
            Case IndustryType.T3Invention
                TempSettings.Facility = DefaultT3InventionFacility
                TempSettings.FacilityType = DefaultT3InventionFacilityType
                TempSettings.ActivityID = IndustryActivities.Invention
                TempSettings.ProductionType = IndustryType.T3Invention
                InventionFacilitySettings = TempSettings
            Case IndustryType.NoPOSManufacturing
                TempSettings.Facility = DefaultManufacturingFacility
                TempSettings.FacilityType = DefaultManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.NoPOSManufacturing
                NoPoSFacilitySettings = TempSettings
            Case IndustryType.POSFuelBlockManufacturing
                TempSettings.Facility = DefaultPOSFuelBlockManufacturingFacility
                TempSettings.FacilityType = DefaultPOSFuelBlockManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.POSFuelBlockManufacturing
                POSFuelBlockFacilitySettings = TempSettings
            Case IndustryType.POSLargeShipManufacturing
                TempSettings.Facility = DefaultPOSLargeShipManufacturingFacility
                TempSettings.FacilityType = DefaultPOSLargeShipManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.POSLargeShipManufacturing
                POSLargeShipFacilitySettings = TempSettings
            Case IndustryType.POSModuleManufacturing
                TempSettings.Facility = DefaultPOSModuleManufacturingFacility
                TempSettings.FacilityType = DefaultPOSModuleManufacturingFacilityType
                TempSettings.ActivityID = IndustryActivities.Manufacturing
                TempSettings.ProductionType = IndustryType.POSModuleManufacturing
                POSModuleFacilitySettings = TempSettings
        End Select

        Return TempSettings

    End Function

#End Region

End Class

' For general program settings
Public Structure ApplicationSettings
    Dim CheckforUpdatesonStart As Boolean
    Dim DataExportFormat As String
    Dim AllowSkillOverride As Boolean
    Dim ShowToolTips As Boolean

    Dim RefiningImplantValue As Double
    Dim ManufacturingImplantValue As Double
    Dim CopyImplantValue As Double

    Dim LoadAssetsonStartup As Boolean
    Dim LoadBPsonStartup As Boolean
    Dim LoadCRESTTeamDataonStartup As Boolean
    Dim LoadCRESTMarketDataonStartup As Boolean
    Dim LoadCRESTFacilityDataonStartup As Boolean
    Dim DisableSound As Boolean

    ' Station Standings for building and selling
    Dim BrokerCorpStanding As Double
    Dim BrokerFactionStanding As Double
    Dim RefineCorpStanding As Double
    Dim RefiningEfficiency As Double ' The default base equipment refining

    Dim RefiningTax As Double ' Tax on refining in stations

    ' ME/TE for BP's we don't own or haven't entered info for
    Dim DefaultBPME As Integer
    Dim DefaultBPTE As Integer

    ' For Build/Buy 
    Dim CheckBuildBuy As Boolean ' Default for setting the check box for build/buy on the blueprints screen
    Dim SuggestBuildBPNotOwned As Boolean ' For Build/Buy suggestions

    Dim LinkBPTabtoFacilitySystem As Boolean ' Links the team drop downs to only teams or auctions in the selected facility system

    Dim DisableSVR As Boolean ' For disabling SVR updates

    ' For shopping list
    Dim ShopListIncludeInventMats As Boolean
    Dim ShopListIncludeREMats As Boolean

    ' The interval for allowing refresh of prices from EVE Central - no less than 1 hour
    Dim EVECentralRefreshInterval As Integer

End Structure

' For BP Tab Settings
Public Structure BPTabSettings
    ' Form stuff
    Dim BlueprintTypeSelection As String
    Dim Tech1Check As Boolean
    Dim Tech2Check As Boolean
    Dim Tech3Check As Boolean
    Dim TechStorylineCheck As Boolean
    Dim TechFactionCheck As Boolean
    Dim TechPirateCheck As Boolean

    Dim SmallCheck As Boolean
    Dim MediumCheck As Boolean
    Dim LargeCheck As Boolean
    Dim XLCheck As Boolean

    Dim IncludeFees As Boolean
    Dim IncludeUsage As Boolean
    Dim IncludeTaxes As Boolean

    Dim IncludeInventionCost As Boolean
    Dim IncludeInventionTime As Boolean
    Dim IncludeCopyCost As Boolean
    Dim IncludeCopyTime As Boolean
    Dim IncludeT3Cost As Boolean
    Dim IncludeT3Time As Boolean

    Dim PricePerUnit As Boolean

    Dim ProductionLines As Integer
    Dim LaboratoryLines As Integer
    Dim T3Lines As Integer

End Structure

' For Update Price Settings
Public Structure UpdatePriceTabSettings
    Dim AllRawMats As Boolean
    Dim Minerals As Boolean
    Dim IceProducts As Boolean
    Dim Gas As Boolean
    Dim Misc As Boolean
    Dim AncientRelics As Boolean
    Dim AncientSalvage As Boolean
    Dim Salvage As Boolean
    Dim Planetary As Boolean
    Dim Datacores As Boolean
    Dim Decryptors As Boolean
    Dim RawMats As Boolean
    Dim ProcessedMats As Boolean
    Dim AdvancedMats As Boolean
    Dim MatsandCompounds As Boolean
    Dim DroneComponents As Boolean
    Dim BoosterMats As Boolean
    Dim Polymers As Boolean
    Dim Asteroids As Boolean

    Dim AllManufacturedItems As Boolean
    Dim Ships As Boolean
    Dim Charges As Boolean

    Dim Modules As Boolean
    Dim Drones As Boolean
    Dim Rigs As Boolean

    Dim Deployables As Boolean
    Dim Subsystems As Boolean
    Dim Boosters As Boolean

    Dim Structures As Boolean
    Dim Celestials As Boolean
    Dim StationComponents As Boolean

    Dim Tools As Boolean
    Dim DataInterfaces As Boolean
    Dim FuelBlocks As Boolean
    Dim Implants As Boolean

    Dim CapT2Components As Boolean
    Dim CapitalComponents As Boolean
    Dim Components As Boolean
    Dim Hybrid As Boolean

    Dim T1 As Boolean
    Dim T2 As Boolean
    Dim T3 As Boolean
    Dim Faction As Boolean
    Dim Pirate As Boolean
    Dim Storyline As Boolean

    Dim SelectedRegions As List(Of String) ' Could have several
    Dim SelectedSystem As String

    Dim PriceImportType As String
    Dim ItemsCombo As String
    Dim RawMatsCombo As String

End Structure

' For Manufacturing Tab Settings
Public Structure ManufacturingTabSettings
    Dim BlueprintType As String

    Dim CheckTech1 As Boolean
    Dim CheckTech2 As Boolean
    Dim CheckTech3 As Boolean
    Dim CheckTechStoryline As Boolean
    Dim CheckTechNavy As Boolean
    Dim CheckTechPirate As Boolean

    Dim ItemTypeFilter As String
    Dim TextItemFilter As String

    Dim CheckBPTypeShips As Boolean
    Dim CheckBPTypeDrones As Boolean
    Dim CheckBPTypeComponents As Boolean
    Dim CheckBPTypeStructures As Boolean
    Dim CheckBPTypeMisc As Boolean
    Dim CheckBPTypeModules As Boolean
    Dim CheckBPTypeAmmoCharges As Boolean
    Dim CheckBPTypeRigs As Boolean
    Dim CheckBPTypeSubsystems As Boolean
    Dim CheckBPTypeBoosters As Boolean
    Dim CheckBPTypeDeployables As Boolean
    Dim CheckBPTypeCelestials As Boolean
    Dim CheckBPTypeStationParts As Boolean

    Dim CheckCapitalComponentsFacility As Boolean
    Dim CheckT3DestroyerFacility As Boolean

    Dim AveragePriceDuration As String

    Dim CheckDecryptorNone As Boolean
    Dim CheckDecryptor06 As Boolean
    Dim CheckDecryptor09 As Boolean
    Dim CheckDecryptor10 As Boolean
    Dim CheckDecryptor11 As Boolean
    Dim CheckDecryptor12 As Boolean
    Dim CheckDecryptor15 As Boolean
    Dim CheckDecryptor18 As Boolean
    Dim CheckDecryptor19 As Boolean

    Dim CheckDecryptorUseforT2 As Boolean
    Dim CheckDecryptorUseforT3 As Boolean

    Dim CheckIgnoreInventionRE As Boolean

    Dim CheckRelicWrecked As Boolean
    Dim CheckRelicIntact As Boolean
    Dim CheckRelicMalfunction As Boolean

    Dim CheckOnlyBuild As Boolean
    Dim CheckOnlyInvent As Boolean
    Dim CheckOnlyRE As Boolean

    Dim CheckIncludeTaxes As Boolean
    Dim CheckIncludeBrokersFees As Boolean
    Dim CheckIncludeUsage As Boolean

    Dim CheckRaceAmarr As Boolean
    Dim CheckRaceCaldari As Boolean
    Dim CheckRaceGallente As Boolean
    Dim CheckRaceMinmatar As Boolean
    Dim CheckRacePirate As Boolean
    Dim CheckRaceOther As Boolean

    Dim SortBy As String

    Dim PriceCompare As String

    Dim CheckIncludeT2Owned As Boolean
    Dim CheckIncludeT3Owned As Boolean

    Dim IgnoreSVRThreshold As Double
    Dim CheckSVRIncludeNull As Boolean

    Dim AveragePriceRegion As String

    Dim ProductionLines As Integer
    Dim LaboratoryLines As Integer
    Dim Runs As Integer

    Dim CheckSmall As Boolean
    Dim CheckMedium As Boolean
    Dim CheckLarge As Boolean
    Dim CheckXL As Boolean

End Structure

' For Datacore Tab Settings
Public Structure DataCoreTabSettings
    Dim PricesFrom As String

    Dim CheckHighSecAgents As Boolean
    Dim CheckLowNullSecAgents As Boolean
    Dim CheckIncludeAgentsCannotAccess As Boolean

    Dim AgentsInRegion As String

    Dim CheckSovAmarr As Boolean
    Dim CheckSovAmmatar As Boolean
    Dim CheckSovGallente As Boolean
    Dim CheckSovSyndicate As Boolean
    Dim CheckSovKhanid As Boolean
    Dim CheckSovThukker As Boolean
    Dim CheckSovCaldari As Boolean
    Dim CheckSovMinmatar As Boolean

    Dim SkillsChecked() As Integer
    Dim SkillsLevel() As Integer

    Dim CorpsChecked() As Integer
    Dim CorpsStanding() As Double

    Dim Connections As Integer
    Dim Negotiation As Integer
    Dim ResearchProjectMgt As Integer

End Structure

' For Reaction Tab Settings
Public Structure ReactionsTabSettings
    Dim POSFuelCost As Double
    Dim NumberofPOS As Integer

    Dim CheckTaxes As Boolean
    Dim CheckFees As Boolean

    Dim CheckAdvMoonMats As Boolean
    Dim CheckProcessedMoonMats As Boolean
    Dim CheckHybrid As Boolean
    Dim CheckComplexBio As Boolean
    Dim CheckSimpleBio As Boolean

    Dim CheckBuildBasic As Boolean
    Dim CheckIgnoreMarket As Boolean
    Dim CheckRefine As Boolean

End Structure

' For Mining Settings
Public Structure MiningTabSettings
    Dim OreType As String ' Ore or Ice

    Dim CheckHighYieldOres As Boolean
    Dim CheckHighSecOres As Boolean
    Dim CheckLowSecOres As Boolean
    Dim CheckNullSecOres As Boolean

    Dim CheckSovAmarr As Boolean
    Dim CheckSovCaldari As Boolean
    Dim CheckSovGallente As Boolean
    Dim CheckSovMinmatar As Boolean
    Dim CheckSovWormhole As Boolean
    Dim CheckSovC1 As Boolean
    Dim CheckSovC2 As Boolean
    Dim CheckSovC3 As Boolean
    Dim CheckSovC4 As Boolean
    Dim CheckSovC5 As Boolean
    Dim CheckSovC6 As Boolean

    Dim CheckIncludeFees As Boolean
    Dim CheckIncludeTaxes As Boolean

    Dim CheckIncludeJumpFuelCosts As Boolean
    Dim TotalJumpFuelCost As Double
    Dim TotalJumpFuelM3 As Double
    Dim JumpCompressedOre As Boolean
    Dim JumpMinerals As Boolean

    Dim OreMiningShip As String
    Dim IceMiningShip As String
    Dim GasMiningShip As String
    Dim OreStrip As String
    Dim IceStrip As String
    Dim GasHarvester As String
    Dim NumOreMiners As Integer
    Dim NumIceMiners As Integer
    Dim NumGasHarvesters As Integer
    Dim OreUpgrade As String
    Dim IceUpgrade As String
    Dim GasUpgrade As String
    Dim NumOreUpgrades As Integer
    Dim NumIceUpgrades As Integer
    Dim NumGasUpgrades As Integer
    Dim OreImplant As String
    Dim IceImplant As String
    Dim GasImplant As String

    Dim MichiiImplant As Boolean
    Dim T2Crystals As Boolean

    Dim CheckUseHauler As Boolean
    Dim RoundTripMin As Integer
    Dim RoundTripSec As Integer
    Dim Haulerm3 As Double

    Dim CheckUseFleetBooster As Boolean
    Dim BoosterShip As String
    Dim BoosterShipSkill As Integer
    Dim MiningFormanSkill As Integer
    Dim MiningDirectorSkill As Integer
    Dim WarfareLinkSpecSkill As Integer
    Dim CheckMineForemanLaserOpBoost As Integer ' 0,1,2
    Dim CheckMineForemanLaserRangeBoost As Integer '0,1,2
    Dim CheckMiningForemanMindLink As Boolean

    Dim CheckRorqDeployed As Boolean
    Dim IndustrialReconfig As Integer

    Dim MiningDroneM3perHour As Double
    Dim RefineOre As Boolean

    Dim MercoxitMiningRig As Boolean
    Dim IceMiningRig As Boolean

End Structure

' For POS Tower Settings
Public Structure PlayerOwnedStationSettings
    ' Form stuff
    Dim TowerRaceID As Integer
    Dim TowerName As String
    Dim CostperHour As Double
    Dim TowerType As String
    Dim TowerSize As String

    Dim FuelBlockBuild As Boolean

    ' Lab numbers
    Dim NumAdvLabs As Integer
    Dim NumMobileLabs As Integer
    Dim NumHyasyodaLabs As Integer

    ' Lab stuff
    Dim MECostperSecond As Double
    Dim TECostperSecond As Double
    Dim InventionCostperSecond As Double
    Dim CopyCostperSecond As Double

    Dim CharterCost As Double

End Structure

' If we show these columns or not
Public Structure IndustryJobsColumnSettings

    ' These are the column orders and shown/not shown. 0 is not shown, else the order number
    Dim JobState As Integer
    Dim TimeToComplete As Integer
    Dim Activity As Integer
    Dim Status As Integer
    Dim StartTime As Integer
    Dim EndTime As Integer
    Dim CompletionTime As Integer
    Dim Blueprint As Integer
    Dim OutputItem As Integer
    Dim OutputItemType As Integer
    Dim InstallSystem As Integer
    Dim InstallRegion As Integer
    Dim LicensedRuns As Integer
    Dim Runs As Integer
    Dim SuccessfulRuns As Integer
    Dim BlueprintLocation As Integer
    Dim OutputLocation As Integer

    Dim JobStateWidth As Integer
    Dim TimeToCompleteWidth As Integer
    Dim ActivityWidth As Integer
    Dim StatusWidth As Integer
    Dim StartTimeWidth As Integer
    Dim EndTimeWidth As Integer
    Dim CompletionTimeWidth As Integer
    Dim BlueprintWidth As Integer
    Dim OutputItemWidth As Integer
    Dim OutputItemTypeWidth As Integer
    Dim InstallSystemWidth As Integer
    Dim InstallRegionWidth As Integer
    Dim LicensedRunsWidth As Integer
    Dim RunsWidth As Integer
    Dim SuccessfulRunsWidth As Integer
    Dim BlueprintLocationWidth As Integer
    Dim OutputLocationWidth As Integer

    Dim OrderByColumn As Integer ' What column index the jobs are sorted
    Dim OrderType As String ' Ascending or Descending

    Dim ViewJobType As String ' Personal, Corp, or Both

    Dim JobTimes As String ' Current or History

End Structure

' If we show these columns or not
Public Structure ManufacturingTabColumnSettings

    ' These are the column orders and shown/not shown. 0 is not shown, else the order number
    Dim ItemCategory As Integer
    Dim ItemGroup As Integer
    Dim ItemName As Integer
    Dim Owned As Integer
    Dim Tech As Integer
    Dim BPME As Integer
    Dim BPTE As Integer
    Dim Inputs As Integer
    Dim Compared As Integer
    Dim Runs As Integer
    Dim ProductionLines As Integer
    Dim LaboratoryLines As Integer
    Dim TotalInventionRECost As Integer
    Dim TotalCopyCost As Integer
    Dim TotalManufacturingCost As Integer
    Dim Taxes As Integer
    Dim BrokerFees As Integer
    Dim BPProductionTime As Integer
    Dim TotalProductionTime As Integer
    Dim ItemMarketPrice As Integer
    Dim Profit As Integer
    Dim ProfitPercentage As Integer
    Dim IskperHour As Integer
    Dim SVR As Integer
    Dim TotalCost As Integer
    Dim BaseJobCost As Integer
    Dim ManufacturingJobFee As Integer
    Dim ManufacturingFacilityName As Integer
    Dim ManufacturingFacilitySystem As Integer
    Dim ManufacturingFacilitySystemIndex As Integer
    Dim ManufacturingFacilityTax As Integer
    Dim ManufacturingFacilityRegion As Integer
    Dim ManufacturingFacilityMEBonus As Integer
    Dim ManufacturingFacilityTEBonus As Integer
    Dim ManufacturingFacilityUsage As Integer
    Dim ComponentFacilityName As Integer
    Dim ComponentFacilitySystem As Integer
    Dim ComponentFacilitySystemIndex As Integer
    Dim ComponentFacilityTax As Integer
    Dim ComponentFacilityRegion As Integer
    Dim ComponentFacilityMEBonus As Integer
    Dim ComponentFacilityTEBonus As Integer
    Dim ComponentFacilityUsage As Integer
    Dim CopyingFacilityName As Integer
    Dim CopyingFacilitySystem As Integer
    Dim CopyingFacilitySystemIndex As Integer
    Dim CopyingFacilityTax As Integer
    Dim CopyingFacilityRegion As Integer
    Dim CopyingFacilityMEBonus As Integer
    Dim CopyingFacilityTEBonus As Integer
    Dim CopyingFacilityUsage As Integer
    Dim InventionREFacilityName As Integer
    Dim InventionREFacilitySystem As Integer
    Dim InventionREFacilitySystemIndex As Integer
    Dim InventionREFacilityTax As Integer
    Dim InventionREFacilityRegion As Integer
    Dim InventionREFacilityMEBonus As Integer
    Dim InventionREFacilityTEBonus As Integer
    Dim InventionREFacilityUsage As Integer
    Dim ManufacturingTeamName As Integer
    Dim ManufacturingTeamBonuses As Integer
    Dim ManufacturingTeamUsage As Integer
    Dim ManufacturingTeamCostModifier As Integer
    Dim ComponentTeamName As Integer
    Dim ComponentTeamBonuses As Integer
    Dim ComponentTeamUsage As Integer
    Dim ComponentTeamCostModifier As Integer
    Dim CopyingTeamName As Integer
    Dim CopyingTeamBonuses As Integer
    Dim CopyingTeamUsage As Integer
    Dim CopyingTeamCostModifier As Integer
    Dim InventionRETeamName As Integer
    Dim InventionRETeamBonuses As Integer
    Dim InventionRETeamUsage As Integer
    Dim InventionRETeamCostModifier As Integer

    Dim ItemCategoryWidth As Integer
    Dim ItemGroupWidth As Integer
    Dim ItemNameWidth As Integer
    Dim OwnedWidth As Integer
    Dim TechWidth As Integer
    Dim BPMEWidth As Integer
    Dim BPTEWidth As Integer
    Dim InputsWidth As Integer
    Dim ComparedWidth As Integer
    Dim RunsWidth As Integer
    Dim ProductionLinesWidth As Integer
    Dim LaboratoryLinesWidth As Integer
    Dim TotalInventionRECostWidth As Integer
    Dim TotalCopyCostWidth As Integer
    Dim TotalManufacturingCostWidth As Integer
    Dim TaxesWidth As Integer
    Dim BrokerFeesWidth As Integer
    Dim BPProductionTimeWidth As Integer
    Dim TotalProductionTimeWidth As Integer
    Dim ItemMarketPriceWidth As Integer
    Dim ProfitWidth As Integer
    Dim ProfitPercentageWidth As Integer
    Dim IskperHourWidth As Integer
    Dim SVRWidth As Integer
    Dim TotalCostWidth As Integer
    Dim BaseJobCostWidth As Integer
    Dim ManufacturingJobFeeWidth As Integer
    Dim ManufacturingFacilityNameWidth As Integer
    Dim ManufacturingFacilitySystemWidth As Integer
    Dim ManufacturingFacilitySystemIndexWidth As Integer
    Dim ManufacturingFacilityTaxWidth As Integer
    Dim ManufacturingFacilityRegionWidth As Integer
    Dim ManufacturingFacilityMEBonusWidth As Integer
    Dim ManufacturingFacilityTEBonusWidth As Integer
    Dim ManufacturingFacilityUsageWidth As Integer
    Dim ComponentFacilityNameWidth As Integer
    Dim ComponentFacilitySystemWidth As Integer
    Dim ComponentFacilitySystemIndexWidth As Integer
    Dim ComponentFacilityTaxWidth As Integer
    Dim ComponentFacilityRegionWidth As Integer
    Dim ComponentFacilityMEBonusWidth As Integer
    Dim ComponentFacilityTEBonusWidth As Integer
    Dim ComponentFacilityUsageWidth As Integer
    Dim CopyingFacilityNameWidth As Integer
    Dim CopyingFacilitySystemWidth As Integer
    Dim CopyingFacilitySystemIndexWidth As Integer
    Dim CopyingFacilityTaxWidth As Integer
    Dim CopyingFacilityRegionWidth As Integer
    Dim CopyingFacilityMEBonusWidth As Integer
    Dim CopyingFacilityTEBonusWidth As Integer
    Dim CopyingFacilityUsageWidth As Integer
    Dim InventionREFacilityNameWidth As Integer
    Dim InventionREFacilitySystemWidth As Integer
    Dim InventionREFacilitySystemIndexWidth As Integer
    Dim InventionREFacilityTaxWidth As Integer
    Dim InventionREFacilityRegionWidth As Integer
    Dim InventionREFacilityMEBonusWidth As Integer
    Dim InventionREFacilityTEBonusWidth As Integer
    Dim InventionREFacilityUsageWidth As Integer
    Dim ManufacturingTeamNameWidth As Integer
    Dim ManufacturingTeamBonusesWidth As Integer
    Dim ManufacturingTeamUsageWidth As Integer
    Dim ManufacturingTeamCostModifierWidth As Integer
    Dim ComponentTeamNameWidth As Integer
    Dim ComponentTeamBonusesWidth As Integer
    Dim ComponentTeamUsageWidth As Integer
    Dim ComponentTeamCostModifierWidth As Integer
    Dim CopyingTeamNameWidth As Integer
    Dim CopyingTeamBonusesWidth As Integer
    Dim CopyingTeamUsageWidth As Integer
    Dim CopyingTeamCostModifierWidth As Integer
    Dim InventionRETeamNameWidth As Integer
    Dim InventionRETeamBonusesWidth As Integer
    Dim InventionRETeamUsageWidth As Integer
    Dim InventionRETeamCostModifierWidth As Integer

    Dim OrderByColumn As Integer ' What column index the jobs are sorted
    Dim OrderType As String ' Ascending or Descending

End Structure

' For Main Industry Flip Belt Settings
Public Structure IndustryFlipBeltSettings
    Dim CycleTime As Double
    Dim m3perCycle As Double
    Dim NumMiners As Integer
    Dim CompressOre As Boolean
    Dim IPHperMiner As Boolean
    Dim IncludeBrokerFees As Boolean
    Dim IncludeTaxes As Boolean
    Dim TrueSec As String
End Structure

' For the checked ore on each mining tab
Public Structure IndustryBeltOreChecks
    Dim Plagioclase As Boolean
    Dim Spodumain As Boolean
    Dim Kernite As Boolean
    Dim Hedbergite As Boolean
    Dim Arkonor As Boolean
    Dim Bistot As Boolean
    Dim Pyroxeres As Boolean
    Dim Crokite As Boolean
    Dim Jaspet As Boolean
    Dim Omber As Boolean
    Dim Scordite As Boolean
    Dim Gneiss As Boolean
    Dim Veldspar As Boolean
    Dim Hemorphite As Boolean
    Dim DarkOchre As Boolean
    Dim Mercoxit As Boolean
    Dim CrimsonArkonor As Boolean
    Dim PrimeArkonor As Boolean
    Dim TriclinicBistot As Boolean
    Dim MonoclinicBistot As Boolean
    Dim SharpCrokite As Boolean
    Dim CrystallineCrokite As Boolean
    Dim OnyxOchre As Boolean
    Dim ObsidianOchre As Boolean
    Dim VitricHedbergite As Boolean
    Dim GlazedHedbergite As Boolean
    Dim VividHemorphite As Boolean
    Dim RadiantHemorphite As Boolean
    Dim PureJaspet As Boolean
    Dim PristineJaspet As Boolean
    Dim LuminousKernite As Boolean
    Dim FieryKernite As Boolean
    Dim AzurePlagioclase As Boolean
    Dim RichPlagioclase As Boolean
    Dim SolidPyroxeres As Boolean
    Dim ViscousPyroxeres As Boolean
    Dim CondensedScordite As Boolean
    Dim MassiveScordite As Boolean
    Dim BrightSpodumain As Boolean
    Dim GleamingSpodumain As Boolean
    Dim ConcentratedVeldspar As Boolean
    Dim DenseVeldspar As Boolean
    Dim IridescentGneiss As Boolean
    Dim PrismaticGneiss As Boolean
    Dim SilveryOmber As Boolean
    Dim GoldenOmber As Boolean
    Dim MagmaMercoxit As Boolean
    Dim VitreousMercoxit As Boolean
End Structure

' For Assets Selected Item Settings
Public Structure AssetWindowSettings

    ' Main window
    Dim AssetType As String
    Dim SortbyName As Boolean

    ' Selected Items
    Dim ItemFilterText As String
    Dim AllItems As Boolean

    Dim AllRawMats As Boolean
    Dim Minerals As Boolean
    Dim IceProducts As Boolean
    Dim Gas As Boolean
    Dim Misc As Boolean
    Dim BPCs As Boolean
    Dim AncientRelics As Boolean
    Dim AncientSalvage As Boolean
    Dim Salvage As Boolean
    Dim Planetary As Boolean
    Dim Datacores As Boolean
    Dim Decryptors As Boolean
    Dim RawMats As Boolean
    Dim ProcessedMats As Boolean
    Dim AdvancedMats As Boolean
    Dim MatsandCompounds As Boolean
    Dim DroneComponents As Boolean
    Dim BoosterMats As Boolean
    Dim Polymers As Boolean
    Dim Asteroids As Boolean

    Dim AllManufacturedItems As Boolean
    Dim Ships As Boolean
    Dim Modules As Boolean
    Dim Drones As Boolean
    Dim Boosters As Boolean
    Dim Rigs As Boolean
    Dim Charges As Boolean
    Dim Subsystems As Boolean
    Dim Structures As Boolean
    Dim Tools As Boolean
    Dim DataInterfaces As Boolean
    Dim CapT2Components As Boolean
    Dim CapitalComponents As Boolean
    Dim Components As Boolean
    Dim Hybrid As Boolean
    Dim FuelBlocks As Boolean
    Dim StationComponents As Boolean
    Dim Celestials As Boolean
    Dim Deployables As Boolean
    Dim Implants As Boolean

    Dim T1 As Boolean
    Dim T2 As Boolean
    Dim T3 As Boolean
    Dim Faction As Boolean
    Dim Pirate As Boolean
    Dim Storyline As Boolean

End Structure

' For the Shopping List
Public Structure ShoppingListSettings
    Dim DataExportFormat As String
    Dim AlwaysonTop As Boolean
    Dim UpdateAssetsWhenUsed As Boolean
    Dim Fees As Boolean
    Dim CalcBuyBuyOrder As Boolean
    Dim Usage As Boolean
    Dim TotalItemTax As Boolean
    Dim TotalItemBrokerFees As Boolean
End Structure

' For all types of facilities
Public Structure FacilitySettings
    Dim Facility As String ' Will be a station/outpost or a pos module
    Dim FacilityType As String ' Type of facility (station, outpost, or pos)
    Dim ActivityID As Integer
    Dim ProductionType As IndustryType ' What will this facility be used for?
    Dim MaterialMultiplier As Double ' May allow them to set the ME/TE for the facility (like in Outposts) when I can't get the info
    Dim TimeMultiplier As Double

    ' For POS, save the location
    Dim SolarSystemID As Long
    Dim SolarSystemName As String
    Dim RegionID As Long
    Dim RegionName As String
    Dim ActivityCostperSecond As Double ' For pos costs
    Dim IncludeActivityCost As Boolean
    Dim IncludeActivityTime As Boolean
    Dim IncludeActivityUsage As Boolean

End Structure

' For saving all 4 possible teams used for jobs
Public Structure TeamSettings
    Dim TeamID As Long
    Dim TeamTab As String ' BP or Calc
End Structure