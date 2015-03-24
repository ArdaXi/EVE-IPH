
Imports System.IO
Imports System.Data.SQLite
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net

' Class for all CREST function calls, which will update tables in the DB if past cache time
Public Class EVECREST
    'Private Const CRESTRootServerURL = "http://public-crest-sisi.testeveonline.com"
    Private Const CRESTRootServerURL = "http://public-crest.eveonline.com"

    ' URL for each file
    Private Const CRESTIndustryTeamSpecialties = "/industry/specialities/"
    Private Const CRESTIndustryTeams = "/industry/teams/"
    Private Const CRESTIndustryTeamAuctions = "/industry/teams/auction/"
    Private Const CRESTIndustrySystems = "/industry/systems/"
    Private Const CRESTIndustryFacilities = "/industry/facilities/"
    Private Const CRESTMarketPrices = "/market/prices/"

    ' Cache field names and times
    Private Const IndustryTeamSpecialtiesCacheDateField As String = "CREST_INDUSTRY_SPECIALIZATIONS_CACHED_UNTIL"
    Private Const IndustryTeamSpecalitiesCacheDateLength As Integer = 1
    Private Const IndustryTeamsCacheDateField As String = "CREST_INDUSTRY_TEAMS_CACHED_UNTIL"
    Private Const IndustryTeamsCacheDateLength As Integer = 1
    Private Const IndustryTeamAuctionsField As String = "CREST_INDUSTRY_TEAM_AUCTIONS_CACHED_UNTIL"
    Private Const IndustryTeamAuctionsLength As Integer = 1
    Private Const IndustrySystemsField As String = "CREST_INDUSTRY_SYSTEMS_CACHED_UNTIL"
    Private Const IndustrySystemsLength As Integer = 1
    Private Const IndustryFacilitiesField As String = "CREST_INDUSTRY_FACILITIES_CACHED_UNTIL"
    Private Const IndustryFacilitiesLength As Integer = 1
    Private Const MarketPricesField As String = "CREST_MARKET_PRICES_CACHED_UNTIL"
    Private Const MarketPricesLength As Integer = 23

    ' market/{region_id}/types/{type_id}/history/ (cache: 23 hours)
    'https://public-crest.eveonline.com/market/10000002/types/34/history/
    ' Provides per day summary of market activity for 13 months for the region_id and type_id sent.

    ' Gets the CREST file from CCP for current Market History and updates the EVEIPH DB with the values
    Public Sub UpdateMarketHistory(ByVal TypeID As Long, RegionID As Long)
        Dim MarketPricesOutput As MarketHistory
        Dim SQL As String
        Dim rsCache As SQLiteDataReader
        Dim CacheDate As Date

        ' First look up the cache date to see if it's time to run the update
        SQL = "SELECT CACHE_DATE FROM MARKET_HISTORY_UPDATE_CACHE WHERE TYPE_ID = " & CStr(TypeID) & " AND REGION_ID = " & CStr(RegionID)
        DBCommand = New SQLiteCommand(SQL, DB)
        rsCache = DBCommand.ExecuteReader

        If rsCache.Read Then
            If Not IsDBNull(rsCache.GetValue(0)) Then
                If rsCache.GetString(0) = "" Then
                    CacheDate = NoDate
                Else
                    CacheDate = CDate(rsCache.GetString(0))
                End If
            Else
                CacheDate = NoDate
            End If
        Else
            CacheDate = NoDate
        End If

        rsCache.Close()
        rsCache = Nothing
        DBCommand = Nothing

        ' If it's later than now, update
        If CacheDate <= Now Then

            Application.DoEvents()

            ' Dump the file into the Specializations object
            MarketPricesOutput = JsonConvert.DeserializeObject(Of MarketHistory) _
                (GetJSONFile(CRESTRootServerURL & "/market/" & CStr(RegionID) & "/types/" & CStr(TypeID) & "/history/", CacheDate))

            ' Read in the data
            If Not IsNothing(MarketPricesOutput) Then
                If MarketPricesOutput.items.Count > 0 Then
                    Call BeginSQLiteTransaction()

                    ' Clear the old record first
                    Call ExecuteNonQuerySQL("DELETE FROM MARKET_HISTORY WHERE TYPE_ID = " & CStr(TypeID) & " AND REGION_ID = " & CStr(RegionID))

                    Application.DoEvents()

                    ' Now read through all the output items and update them in ITEM_PRICES
                    For i = 0 To MarketPricesOutput.totalCount - 1
                        With MarketPricesOutput.items(i)
                            SQL = "INSERT INTO MARKET_HISTORY VALUES (" & CStr(TypeID) & "," & CStr(RegionID) & "," & CStr(.volume) & ","
                            SQL = SQL & CStr(.lowPrice) & "," & CStr(.highPrice) & "," & CStr(.avgPrice) & "," & CStr(.orderCount) & ",'" & .date_str.Replace("T", " ") & "')"
                            Call ExecuteNonQuerySQL(SQL)
                        End With

                        Application.DoEvents()
                    Next

                    ' Set the Cache Date to now plus the length since it's not sent in the file
                    Call ExecuteNonQuerySQL("DELETE FROM MARKET_HISTORY_UPDATE_CACHE WHERE TYPE_ID = " & CStr(TypeID) & " AND REGION_ID = " & CStr(RegionID))
                    Call ExecuteNonQuerySQL("INSERT INTO MARKET_HISTORY_UPDATE_CACHE VALUES (" & CStr(TypeID) & "," & CStr(RegionID) & "," & "'" & CacheDate & "')")

                    ' Done updating
                    Call CommitSQLiteTransaction()
                End If
            End If
        End If

    End Sub

    ' For Market History
    Private Class MarketHistory
        '{"totalCount_str": "414", "items": [], "pageCount": 1, "pageCount_str": "1", "totalCount": 414}
        <JsonProperty("totalCount_str")> Public totalCount_str As String
        <JsonProperty("items")> Public items() As MarketPriceItems
        <JsonProperty("pageCount")> Public pageCount As Integer
        <JsonProperty("pageCount_str")> Public pageCount_str As String
        <JsonProperty("totalCount")> Public totalCount As Integer
    End Class

    ' For Market History
    Private Class MarketPriceItems
        '{"volume_str": "28662910175", "orderCount": 4312, "lowPrice": 4.98, "highPrice": 5.04, "avgPrice": 5.0, "volume": 28662910175, "orderCount_str": "4312", "date": "2014-02-01T00:00:00"}
        <JsonProperty("volume_str")> Public volume_str As String
        <JsonProperty("orderCount")> Public orderCount As Long
        <JsonProperty("lowPrice")> Public lowPrice As Double
        <JsonProperty("highPrice")> Public highPrice As Double
        <JsonProperty("avgPrice")> Public avgPrice As Double
        <JsonProperty("volume")> Public volume As Long
        <JsonProperty("orderCount_str")> Public orderCount_str As String
        <JsonProperty("date")> Public date_str As String
    End Class

    ' Current endpoints as of July 22, 2014

    ' /industry/specialities/ (cache: 1 hour)
    ' vnd.ccp.eve.IndustrySpecialityCollection-v1
    ' Lists all specialties that can be associated with teams and what groups those specialties modify.

    ' Gets the CREST file from CCP for current Industry Specialties and updates the EVEIPH DB with the values
    Public Sub UpdateIndustrySpecialties(Optional ByRef UpdateLabel As Label = Nothing, Optional ByRef PB As ProgressBar = Nothing)
        Dim TeamSpecialtiesOutput As TeamSpecializations
        Dim SQL As String

        ' Data Variables
        Dim SpecializationName As String
        Dim SpecializationID As Integer
        Dim GroupID As Integer
        Dim CacheDate As Date

        Dim TempLabel As Label
        Dim TempPB As ProgressBar

        If IsNothing(UpdateLabel) Then
            TempLabel = New Label
        Else
            TempLabel = UpdateLabel
        End If

        If IsNothing(PB) Then
            TempPB = New ProgressBar
        Else
            TempPB = PB
        End If

        ' First look up the cache date to see if it's time to run the update
        CacheDate = GetCRESTCacheDate(IndustryTeamSpecialtiesCacheDateField)

        ' If it's later than now, update
        If CacheDate <= Now Then

            TempLabel.Text = "Downloading Team Specialty Data..."
            Application.DoEvents()

            ' Dump the file into the Specializations object
            TeamSpecialtiesOutput = JsonConvert.DeserializeObject(Of TeamSpecializations) _
                (GetJSONFile(CRESTRootServerURL & CRESTIndustryTeamSpecialties, CacheDate))

            ' Read in the data
            If Not IsNothing(TeamSpecialtiesOutput) Then
                If TeamSpecialtiesOutput.items.Count > 0 Then
                    Call BeginSQLiteTransaction()

                    ' Delete the old records first
                    Call ExecuteNonQuerySQL("DELETE FROM INDUSTRY_GROUP_SPECIALTIES")
                    Call ExecuteNonQuerySQL("DELETE FROM INDUSTRY_CATEGORY_SPECIALTIES")

                    ' Now read through all the output items and input them into the DB
                    For i = 0 To TeamSpecialtiesOutput.totalCount - 1

                        TempLabel.Text = "Saving Team Specialty Data..."
                        TempPB.Minimum = 0
                        TempPB.Value = 0
                        TempPB.Maximum = TeamSpecialtiesOutput.totalCount - 1
                        TempPB.Visible = True
                        Application.DoEvents()

                        SpecializationName = TeamSpecialtiesOutput.items(i).name
                        SpecializationID = TeamSpecialtiesOutput.items(i).id

                        ' Now loop through all the items and get the group IDs that this specialization uses. Group ID is the invGroups ID for linking to items
                        If TeamSpecialtiesOutput.items(i).groups.Count <> 0 Then
                            For j = 0 To TeamSpecialtiesOutput.items(i).groups.Count - 1
                                GroupID = TeamSpecialtiesOutput.items(i).groups(j).id

                                SQL = "INSERT INTO INDUSTRY_GROUP_SPECIALTIES VALUES(" & GroupID & "," & SpecializationID & ",'" & SpecializationName & "')"
                                Call ExecuteNonQuerySQL(SQL)

                            Next
                        Else
                            ' Insert the record with no groups - these are the category of the team, which then can do groups of items in that category
                            SQL = "INSERT INTO INDUSTRY_CATEGORY_SPECIALTIES VALUES(" & SpecializationID & ",'" & SpecializationName & "')"
                            Call ExecuteNonQuerySQL(SQL)
                        End If

                        TempPB.Value = i
                        Application.DoEvents()

                    Next

                    ' Rebuild indexes
                    Call ExecuteNonQuerySQL("REINDEX IDX_CAT_ID")

                    ' Set the Cache Date to now plus the length since it's not sent in the file
                    Call SetCRESTCacheDate(IndustryTeamSpecialtiesCacheDateField, DateAdd("h", IndustryTeamSpecalitiesCacheDateLength, Now))
                    ' Done updating
                    Call CommitSQLiteTransaction()
                End If
            End If
        End If

    End Sub

    ' For Team Specalizations
    Private Class TeamSpecializations

        '"{"totalCount_str": "141", "items": [], "pageCount": 1, "pageCount_str": "1", "totalCount": 141}
        <JsonProperty("totalCount_str")> Public totalCount_str As String
        <JsonProperty("items")> Public items() As TeamSpecializationsItem
        <JsonProperty("pageCount")> Public pageCount As Integer
        <JsonProperty("pageCount_str")> Public pageCount_str As String
        <JsonProperty("totalCount")> Public totalCount As Integer

    End Class

    ' For Team Specalizations
    Private Class TeamSpecializationsItem
        '{"id_str": "0", "id": 0, "groups": [], "name": "Specializations"}
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("id")> Public id As Integer
        <JsonProperty("groups")> Public groups() As TeamSpecializationsGroupIDs
        <JsonProperty("name")> Public name As String
    End Class

    ' For Team Specalizations
    Private Class TeamSpecializationsGroupIDs
        '{"id": 838, "id_str": "838"}, 
        <JsonProperty("id")> Public id As Integer
        <JsonProperty("id_str")> Public id_str As String
    End Class

    ' /industry/teams/ (cache: 1 hour)
    ' vnd.ccp.eve.IndustryTeamCollection-v1
    ' Returns a list of all active teams, excluding the teams in auction. This does not include wormhole space.

    ' Gets the CREST file from CCP for current Industry Teams and updates the EVEIPH DB with the values
    Public Sub UpdateIndustryTeams(Optional ByRef UpdateLabel As Label = Nothing, Optional ByRef PB As ProgressBar = Nothing)
        Dim IndustryTeamsOutput As IndustryTeams
        Dim SQL As String

        ' Data Variables
        Dim CacheDate As Date
        ' Industry Teams
        Dim TeamID As String
        Dim TeamName As String
        Dim ActivityID As String
        Dim CreationTime As DateTime
        Dim ExpiryTime As DateTime
        Dim CostModifier As Double
        Dim SolarSystemID As Long
        Dim SolarSystemName As String
        Dim CategorySpecializationID As Integer

        ' Industry Teams Bonuses
        Dim BonusID As Integer
        Dim BonusType As String
        Dim BonusValue As Double
        Dim GroupSpecializationID As Integer

        Dim rsCheck As SQLiteDataReader
        Dim rsCheck2 As SQLiteDataReader
        Dim Sequence As Integer = 0

        Dim TempLabel As Label
        Dim TempPB As ProgressBar

        If IsNothing(UpdateLabel) Then
            TempLabel = New Label
        Else
            TempLabel = UpdateLabel
        End If

        If IsNothing(PB) Then
            TempPB = New ProgressBar
        Else
            TempPB = PB
        End If

        ' First look up the cache date to see if it's time to run the update
        CacheDate = GetCRESTCacheDate(IndustryTeamsCacheDateField)

        ' If it's later than now, update
        If CacheDate <= Now Then

            TempLabel.Text = "Downloading Team Data..."
            Application.DoEvents()

            ' Dump the file into the Specializations object
            IndustryTeamsOutput = JsonConvert.DeserializeObject(Of IndustryTeams)(GetJSONFile(CRESTRootServerURL & CRESTIndustryTeams, CacheDate))

            ' Read in the data
            If Not IsNothing(IndustryTeamsOutput) Then
                If IndustryTeamsOutput.items.Count > 0 Then
                    Call BeginSQLiteTransaction()

                    ' Delete the old records first
                    Call ExecuteNonQuerySQL("DELETE FROM INDUSTRY_TEAMS")
                    Call ExecuteNonQuerySQL("DELETE FROM INDUSTRY_TEAMS_BONUSES")

                    ' Now read through all the output items and input them into the DB
                    For i = 0 To IndustryTeamsOutput.totalCount - 1

                        TempLabel.Text = "Saving Team Data..."
                        TempPB.Minimum = 0
                        TempPB.Value = 0
                        TempPB.Maximum = IndustryTeamsOutput.totalCount - 1
                        TempPB.Visible = True
                        Application.DoEvents()

                        With IndustryTeamsOutput.items(i)
                            TeamID = .id
                            TeamName = FormatDBString(Replace(.name, "<br>", " ")) ' Replace the break with a space
                            ActivityID = .activity
                            CreationTime = CDate(Replace(.creationTime, "T", " "))
                            ExpiryTime = CDate(Replace(.expiryTime, "T", " "))
                            CostModifier = .costModifier
                            SolarSystemID = .solarSystem.id
                            SolarSystemName = .solarSystem.name
                            CategorySpecializationID = .specialization.id
                        End With

                        ' Insert this team record
                        SQL = "INSERT INTO INDUSTRY_TEAMS VALUES(" & TeamID & ",'" & TeamName & "'," & ActivityID & ","
                        SQL = SQL & CStr(SolarSystemID) & ",'" & SolarSystemName & "',"
                        SQL = SQL & CStr(CostModifier) & ",'"
                        SQL = SQL & Format(CreationTime, SQLiteDateFormat) & "','"
                        SQL = SQL & Format(ExpiryTime, SQLiteDateFormat) & "',"
                        SQL = SQL & CStr(CategorySpecializationID) & ")"

                        Call ExecuteNonQuerySQL(SQL)

                        ' Now loop through all the workers and input the bonus data
                        For j = 0 To IndustryTeamsOutput.items(i).workers.Count - 1

                            With IndustryTeamsOutput.items(i).workers(j)
                                BonusID = .bonus.id
                                BonusType = .bonus.bonusType
                                BonusValue = .bonus.value
                                GroupSpecializationID = .specialization.id
                            End With

                            SQL = "INSERT INTO INDUSTRY_TEAMS_BONUSES VALUES (" & TeamID & ",'" & TeamName & "'," & CStr(BonusID) & ",'"
                            SQL = SQL & BonusType & "'," & CStr(-1 * BonusValue) & "," & CStr(GroupSpecializationID) & ")" ' make bonus value positive

                            Call ExecuteNonQuerySQL(SQL)

                        Next

                        TempPB.Value = i
                        Application.DoEvents()
                    Next

                    ' Set the Cache Date to now plus the length since it's not sent in the file
                    Call SetCRESTCacheDate(IndustryTeamsCacheDateField, CacheDate)

                    TempLabel.Text = "Finalizing Team Data..."
                    TempPB.Visible = False
                    Application.DoEvents()

                    ' Finally, update the team names that are duplicates - short term fix for duplicates?
                    SQL = "SELECT TEAM_NAME, COUNT(*) FROM INDUSTRY_TEAMS GROUP BY TEAM_NAME HAVING COUNT(*) > 1"
                    DBCommand = New SQLiteCommand(SQL, DB)
                    rsCheck = DBCommand.ExecuteReader

                    While rsCheck.Read
                        ' Update each team name to make it unique
                        SQL = "SELECT TEAM_NAME, TEAM_ID FROM INDUSTRY_TEAMS WHERE TEAM_NAME = '" & rsCheck.GetString(0) & "' ORDER BY TEAM_ID"
                        DBCommand = New SQLiteCommand(SQL, DB)
                        rsCheck2 = DBCommand.ExecuteReader

                        While rsCheck2.Read
                            ' Update Teams first
                            SQL = "UPDATE INDUSTRY_TEAMS SET TEAM_NAME = '" & FormatDBString(rsCheck2.GetString(0) & " " & CStr(Sequence + 1)) & "' "
                            SQL = SQL & "WHERE TEAM_NAME = '" & rsCheck2.GetString(0) & "' AND TEAM_ID = " & rsCheck2.GetInt64(1)
                            Call ExecuteNonQuerySQL(SQL)

                            ' Update Team Bonuses 
                            SQL = "UPDATE INDUSTRY_TEAMS_BONUSES SET TEAM_NAME = '" & FormatDBString(rsCheck2.GetString(0) & " " & CStr(Sequence + 1)) & "' "
                            SQL = SQL & "WHERE TEAM_NAME = '" & rsCheck2.GetString(0) & "' AND TEAM_ID = " & rsCheck2.GetInt64(1)
                            Call ExecuteNonQuerySQL(SQL)

                            ' Increment
                            Sequence += 1
                        End While

                        rsCheck2.Close()
                        Sequence = 0 ' Reset for next group

                    End While

                    ' Rebuild indexes on teams and bonuses
                    Call ExecuteNonQuerySQL("REINDEX IDX_TEAMS_TEAM_NAME")
                    Call ExecuteNonQuerySQL("REINDEX IDX_TEAMS_ACTIVITY_ID")
                    Call ExecuteNonQuerySQL("REINDEX IDX_TEAMS_TEAM_ID")

                    Call ExecuteNonQuerySQL("REINDEX IDX_BONUSES_ID_GID")
                    Call ExecuteNonQuerySQL("REINDEX IDX_BONUSES_NAME_GID")

                    rsCheck.Close()
                    DBCommand = Nothing
                    rsCheck = Nothing
                    rsCheck2 = Nothing

                    ' Done updating
                    Call CommitSQLiteTransaction()
                End If
            End If
        End If

    End Sub

    ' For Industry Teams
    Private Class IndustryTeams
        '{"totalCount_str": "56", "items": [], "pageCount": 1, "pageCount_str": "1", "totalCount": 56}
        <JsonProperty("totalCount_str")> Public totalCount_str As String
        <JsonProperty("items")> Public items() As IndustryTeamItems
        <JsonProperty("pageCount")> Public pageCount As Integer
        <JsonProperty("pageCount_str")> Public pageCount_str As String
        <JsonProperty("totalCount")> Public totalCount As Integer
    End Class

    ' For Industry Teams
    Private Class IndustryTeamItems
        '{"solarSystem": {}, "name": "Roden Shipyards<br>Team MPV00", "creationTime": "2014-07-11T22:56:52",  "workers": [],
        ' "expiryTime": "2014-08-15T22:56:52", "costModifier": 5.0, "id_str": "22", "activity": 1, "activity_str": "1",
        ' "id": 22},  "specialization": {} "
        <JsonProperty("solarSystem")> Public solarSystem As SolarSystem
        <JsonProperty("name")> Public name As String
        <JsonProperty("creationTime")> Public creationTime As String
        <JsonProperty("workers")> Public workers() As IndustryTeamsWorkers
        <JsonProperty("expiryTime")> Public expiryTime As String
        <JsonProperty("costModifier")> Public costModifier As Double
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("activity")> Public activity As String
        <JsonProperty("activity_str")> Public activity_str As String
        <JsonProperty("id")> Public id As String
        <JsonProperty("specialization")> Public specialization As IndustryTeamsSpecialization
    End Class

    ' For Industry Teams
    Private Class IndustryTeamsWorkers
        '"workers": [{bonus},{specialization}]
        <JsonProperty("bonus")> Public bonus As IndustryTeamsBonus
        <JsonProperty("specialization")> Public specialization As IndustryTeamsSpecialization
    End Class

    ' For Industry Teams
    Private Class IndustryTeamsBonus
        '{"bonus": {"bonusType": "TE", "id": 0, "value": -1.0, "id_str": "0"}, 
        <JsonProperty("bonusType")> Public bonusType As String
        <JsonProperty("id")> Public id As Integer
        <JsonProperty("value")> Public value As Double
        <JsonProperty("id_str")> Public id_str As String
    End Class

    ' For Industry Teams
    Private Class IndustryTeamsSpecialization
        '"specialization": {"href": "http://public-crest-sisi.testeveonline.com/industry/specialities/6/", "id": 6, "id_str": "6"},
        <JsonProperty("href")> Public href As String
        <JsonProperty("id")> Public id As Integer
        <JsonProperty("id_str")> Public id_str As String
    End Class

    ' /industry/teams/auction/ (cache: 1 hour)
    ' vnd.ccp.eve.IndustryTeamCollection-v1
    ' Returns a list of all the teams currently up for auction.

    ' Gets the CREST file from CCP for current Industry Team Auctions and updates the EVEIPH DB with the values
    Public Sub UpdateIndustryTeamAuctions(Optional ByRef UpdateLabel As Label = Nothing, Optional ByRef PB As ProgressBar = Nothing)
        Dim IndustryTeamAuctionsOutput As IndustryTeamAuctions
        Dim SQL As String
        Dim CacheDate As Date
        Dim AuctionID As String
        Dim CreationTime As DateTime
        Dim ExpiryTime As DateTime
        Dim TeamName As String
        Dim TeamID As Long

        Dim rsCheck As SQLiteDataReader
        Dim rsCheck2 As SQLiteDataReader
        Dim Sequence As Integer = 0

        Dim TempLabel As Label
        Dim TempPB As ProgressBar

        If IsNothing(UpdateLabel) Then
            TempLabel = New Label
        Else
            TempLabel = UpdateLabel
        End If

        If IsNothing(PB) Then
            TempPB = New ProgressBar
        Else
            TempPB = PB
        End If

        ' First look up the cache date to see if it's time to run the update
        CacheDate = GetCRESTCacheDate(IndustryTeamAuctionsField)

        ' If it's later than now, update
        If CacheDate <= Now Then

            TempLabel.Text = "Downloading Team Auction Data..."
            Application.DoEvents()

            ' Dump the file into the Specializations object
            IndustryTeamAuctionsOutput = JsonConvert.DeserializeObject(Of IndustryTeamAuctions)(GetJSONFile(CRESTRootServerURL & CRESTIndustryTeamAuctions, CacheDate))

            ' Read in the data
            If Not IsNothing(IndustryTeamAuctionsOutput) Then
                If IndustryTeamAuctionsOutput.items.Count > 0 Then
                    Call BeginSQLiteTransaction()

                    TempLabel.Text = "Saving Team Auction Data..."
                    TempPB.Minimum = 0
                    TempPB.Value = 0
                    TempPB.Maximum = IndustryTeamAuctionsOutput.totalCount - 1
                    TempPB.Visible = True
                    Application.DoEvents()

                    ' Delete the old records first
                    Call ExecuteNonQuerySQL("DELETE FROM INDUSTRY_TEAMS_AUCTIONS")

                    ' Now read through all the output items and input them into the DB
                    For i = 0 To IndustryTeamAuctionsOutput.totalCount - 1

                        With IndustryTeamAuctionsOutput.items(i)
                            AuctionID = CStr(.id)
                            CreationTime = CDate(Replace(.creationTime, "T", " "))
                            ExpiryTime = CDate(Replace(.expiryTime, "T", " "))
                            TeamName = FormatDBString(Replace(.name, "<br>", " "))
                            TeamID = CLng(.id)

                            ' Insert the base auction data 
                            SQL = "INSERT INTO INDUSTRY_TEAMS_AUCTIONS VALUES(" & TeamID & ",'" & TeamName & "'," & CStr(.activity) & ","
                            SQL = SQL & CStr(.solarSystem.id) & ",'" & .solarSystem.name & "'," & CStr(.costModifier) & ",'"
                            SQL = SQL & Format(CreationTime, SQLiteDateFormat) & "','" & Format(ExpiryTime, SQLiteDateFormat) & "',"
                            SQL = SQL & CStr(.specialization.id) & "," & AuctionID & ")"

                            Call ExecuteNonQuerySQL(SQL)

                            For j = 0 To .workers.Count - 1
                                ' Insert the workers for this team and link by team name (should be unique) - make bonus value negative
                                SQL = "INSERT INTO INDUSTRY_TEAMS_BONUSES VALUES (" & TeamID & ",'" & TeamName & "'," & CStr(.workers(j).bonus.id) & ",'"
                                SQL = SQL & .workers(j).bonus.bonusType & "'," & CStr(-1 * .workers(j).bonus.value) & "," & CStr(.workers(j).specialization.id) & ")"

                                Call ExecuteNonQuerySQL(SQL)
                            Next

                            ' Insert the Bid data
                            For j = 0 To .solarSystemBids.Count - 1
                                ' For each bid in the array of solar systems, you can have multiple bids, and each bid can have multiple characters
                                For k = 0 To .solarSystemBids(j).characterBids.Count - 1

                                    SQL = "INSERT INTO INDUSTRY_TEAMS_AUCTIONS_BIDS VALUES(" & AuctionID & ","
                                    SQL = SQL & CStr(.solarSystemBids(j).solarSystem.id) & ",'"
                                    SQL = SQL & .solarSystemBids(j).solarSystem.name & "',"
                                    SQL = SQL & CStr(.solarSystemBids(j).bidAmount) & ","
                                    SQL = SQL & CStr(.solarSystemBids(j).characterBids(k).character.CharacterID) & ",'"
                                    SQL = SQL & FormatDBString(.solarSystemBids(j).characterBids(k).character.name) & "',"
                                    SQL = SQL & CStr(CInt(.solarSystemBids(j).characterBids(k).character.isNPC)) & ","
                                    SQL = SQL & CStr(.solarSystemBids(j).characterBids(k).bidAmount) & ")"

                                    Call ExecuteNonQuerySQL(SQL)

                                Next
                            Next
                        End With

                        TempPB.Value = i
                        Application.DoEvents()
                    Next

                    ' Set the Cache Date to now plus the length since it's not sent in the file
                    Call SetCRESTCacheDate(IndustryTeamAuctionsField, CacheDate)

                    TempLabel.Text = "Finalizing Team Auction Data..."
                    TempPB.Visible = False
                    Application.DoEvents()

                    ' Finally, update the team names that are duplicates - short term fix for duplicates?
                    SQL = "SELECT TEAM_NAME, COUNT(*) FROM INDUSTRY_TEAMS_AUCTIONS GROUP BY TEAM_NAME HAVING COUNT(*) >1"
                    DBCommand = New SQLiteCommand(SQL, DB)
                    rsCheck = DBCommand.ExecuteReader

                    While rsCheck.Read
                        ' Update each team name to make it unique
                        SQL = "SELECT TEAM_NAME, TEAM_ID FROM INDUSTRY_TEAMS_AUCTIONS WHERE TEAM_NAME = '" & rsCheck.GetString(0) & "' ORDER BY TEAM_ID"
                        DBCommand = New SQLiteCommand(SQL, DB)
                        rsCheck2 = DBCommand.ExecuteReader

                        While rsCheck2.Read
                            ' Update Teams first
                            SQL = "UPDATE INDUSTRY_TEAMS_AUCTIONS SET TEAM_NAME = '" & FormatDBString(rsCheck2.GetString(0) & " " & CStr(Sequence + 1)) & "' "
                            SQL = SQL & "WHERE TEAM_NAME = '" & rsCheck2.GetString(0) & "' AND TEAM_ID = " & rsCheck2.GetInt64(1)
                            Call ExecuteNonQuerySQL(SQL)

                            ' Update Team Bonuses 
                            SQL = "UPDATE INDUSTRY_TEAMS_BONUSES SET TEAM_NAME = '" & FormatDBString(rsCheck2.GetString(0) & " " & CStr(Sequence + 1)) & "' "
                            SQL = SQL & "WHERE TEAM_NAME = '" & rsCheck2.GetString(0) & "' AND TEAM_ID = " & rsCheck2.GetInt64(1)
                            Call ExecuteNonQuerySQL(SQL)
                            ' Increment
                            Sequence += 1
                        End While

                        rsCheck2.Close()
                        Sequence = 0 ' Reset for next group

                    End While

                    ' Rebuild indexes on auctions and bonuses
                    Call ExecuteNonQuerySQL("REINDEX IDX_AUCTIONS_TEAM_ID")
                    Call ExecuteNonQuerySQL("REINDEX IDX_AUCTIONS_CTIVITY_ID")
                    Call ExecuteNonQuerySQL("REINDEX IDX_AUCTIONS_TEAM_NAME")
                    Call ExecuteNonQuerySQL("REINDEX IDX_AUCTIONS_AUCTION_ID")
                    Call ExecuteNonQuerySQL("REINDEX IDX_BIDS_AUCTION_ID")

                    Call ExecuteNonQuerySQL("REINDEX IDX_BONUSES_ID_GID")
                    Call ExecuteNonQuerySQL("REINDEX IDX_BONUSES_NAME_GID")

                    rsCheck.Close()
                    DBCommand = Nothing
                    rsCheck = Nothing
                    rsCheck2 = Nothing

                    ' Done updating
                    Call CommitSQLiteTransaction()

                End If
            End If
        End If

    End Sub

    ' For Industry Team Auctions
    Private Class IndustryTeamAuctions
        '{"totalCount_str": "1005", "items":[], "pageCount": 1, "pageCount_str": "1", "totalCount": 1005}
        <JsonProperty("totalCount_str")> Public totalCount_str As String
        <JsonProperty("items")> Public items() As IndustryTeamAuction
        <JsonProperty("pageCount")> Public pageCount As Integer
        <JsonProperty("pageCount_str")> Public pageCount_str As String
        <JsonProperty("totalCount")> Public totalCount As Integer
    End Class

    ' For Industry Team Auctions
    Private Class IndustryTeamAuction
        '{"solarSystem": {}, "specialization": {}, "creationTime": "2014-07-13T14:23:38", "workers": [], 
        '"expiryTime": "2014-08-17T14:23:38", "solarSystemBids": [], "auctionExpiryTime": "2014-07-20T14:23:38", 
        '"id_str": "258", "activity": 1, "costModifier": 4.0, "id": 1281, "name": "Guristas Production<br>Team MPV00"},
        <JsonProperty("solarSystem")> Public solarSystem As SolarSystem
        <JsonProperty("specialization")> Public specialization As IndustryTeamsSpecialization
        <JsonProperty("creationTime")> Public creationTime As String
        <JsonProperty("workers")> Public workers() As IndustryTeamsWorkers
        <JsonProperty("expiryTime")> Public expiryTime As String
        <JsonProperty("solarSystemBids")> Public solarSystemBids() As IndustryTeamAuctionSolarSystemBids
        <JsonProperty("auctionExpiryTime")> Public auctionExpiryTime As String
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("activity")> Public activity As String
        <JsonProperty("costModifier")> Public costModifier As Double
        <JsonProperty("id")> Public id As Integer
        <JsonProperty("name")> Public name As String

    End Class

    ' For Industry Team Auctions
    Private Class IndustryTeamAuctionSolarSystemBids
        '"solarSystemBids": [{"characterBids": [], "solarSystem": {}, "bidAmount": 400000.0}], 
        <JsonProperty("characterBids")> Public characterBids() As IndustryTeamAuctionCharacterBid
        <JsonProperty("solarSystem")> Public solarSystem As SolarSystem
        <JsonProperty("bidAmount")> Public bidAmount As Double
    End Class

    ' For Industry Team Auctions
    Private Class IndustryTeamAuctionCharacterBid
        '"characterBids": [{"character": {}, "bidAmount": 400000.0}], 
        <JsonProperty("character")> Public character As IndustryTeamAuctionCharacter
        <JsonProperty("bidAmount")> Public bidAmount As Double
    End Class

    ' For Industry Team Auctions
    Private Class IndustryTeamAuctionCharacter
        ' "character": {"isNPC": false, "id": 94733327, "href": "http://public-crest-sisi.testeveonline.com/characters/94733327/", "name": "Aniv Omegadeleiv", "capsuleer": {"href": ""}
        <JsonProperty("isNPC")> Public isNPC As Boolean
        <JsonProperty("id")> Public CharacterID As Long
        <JsonProperty("href")> Public href As String
        <JsonProperty("name")> Public name As String
        <JsonProperty("capsuleer")> Public capsuleer As hrefKey
    End Class

    ' /industry/facilities/ (cache: 1 hour)
    ' vnd.ccp.eve.IndustryFacilityCollection-v1
    ' This returns a list of all publicly accessible facilities, including player built outposts in nullsec.

    ' Gets the CREST file from CCP for current Industry Facilities and updates the EVEIPH DB with the values
    Public Sub UpdateIndustryFacilties(Optional ByRef UpdateLabel As Label = Nothing, Optional ByRef PB As ProgressBar = Nothing)
        Dim IndustryFacilitiesOutput As IndustryFacilities
        Dim SQL As String
        Dim CacheDate As Date

        Dim TempLabel As Label
        Dim TempPB As ProgressBar

        Dim FacilityName As String

        If IsNothing(UpdateLabel) Then
            TempLabel = New Label
        Else
            TempLabel = UpdateLabel
        End If

        If IsNothing(PB) Then
            TempPB = New ProgressBar
        Else
            TempPB = PB
        End If

        ' First look up the cache date to see if it's time to run the update
        CacheDate = GetCRESTCacheDate(IndustryFacilitiesField)

        ' If it's later than now, update
        If CacheDate <= Now Then

            TempLabel.Text = "Downloading Facility Data..."
            Application.DoEvents()

            ' Dump the file into the Specializations object
            IndustryFacilitiesOutput = JsonConvert.DeserializeObject(Of IndustryFacilities) _
                (GetJSONFile(CRESTRootServerURL & CRESTIndustryFacilities, CacheDate))

            ' Read in the data
            If Not IsNothing(IndustryFacilitiesOutput) Then
                If IndustryFacilitiesOutput.items.Count > 0 Then
                    Call BeginSQLiteTransaction()
                    TempLabel.Text = "Saving Industry Facilities Data..."
                    TempPB.Minimum = 0
                    TempPB.Value = 0
                    TempPB.Maximum = IndustryFacilitiesOutput.totalCount - 1
                    TempPB.Visible = True
                    Application.DoEvents()

                    ' Delete the old records first
                    Call ExecuteNonQuerySQL("DELETE FROM INDUSTRY_FACILITIES")

                    ' Now read through all the output items and input them into the DB
                    For i = 0 To IndustryFacilitiesOutput.totalCount - 1
                        With IndustryFacilitiesOutput.items(i)
                            ' See if this is an outpost or not and add the tag for type to the name
                            Select Case .type.id
                                ' FACILITY_TYPE_ID	FACILITY_TYPE
                                ' 21644	Amarr Factory Outpost
                                ' 21645	Gallente Administrative Outpost
                                ' 21646	Minmatar Service Outpost
                                ' 21642	Caldari Research Outpost
                                Case 21644
                                    FacilityName = Format(.name) & " (A)"
                                Case 21645
                                    FacilityName = Format(.name) & " (G)"
                                Case 21646
                                    FacilityName = Format(.name) & " (M)"
                                Case 21642
                                    FacilityName = Format(.name) & " (C)"
                                Case Else
                                    FacilityName = Format(.name)
                            End Select

                            SQL = "INSERT INTO INDUSTRY_FACILITIES VALUES (" & CStr(.facilityID) & ",'" & FormatDBString(FacilityName) & "'," & CStr(.type.id) & ","
                            SQL = SQL & CStr(.tax) & "," & CStr(.solarSystem.id) & "," & CStr(.region.id) & ","
                            SQL = SQL & CStr(.owner.id) & ")"
                            Call ExecuteNonQuerySQL(SQL)
                        End With

                        TempPB.Value = i
                        Application.DoEvents()
                    Next

                    ' Rebuild indexes on INDUSTRY_FACILITIES
                    Call ExecuteNonQuerySQL("REINDEX IDX_IF_MAIN")
                    Call ExecuteNonQuerySQL("REINDEX IDX_IF_SSID")

                    ' Now that everything is inserted build a master station table that we can query for anything
                    TempLabel.Text = "Building Stations Table..."
                    Application.DoEvents()

                    ' First drop the table if it exists
                    SQL = "DROP TABLE IF EXISTS STATION_FACILITIES"
                    Call ExecuteNonQuerySQL(SQL)

                    ' Run the big query
                    SQL = "CREATE TABLE STATION_FACILITIES AS SELECT * FROM ( "
                    SQL = SQL & "SELECT FACILITY_ID, FACILITY_NAME, "
                    SQL = SQL & "SOLAR_SYSTEMS.solarSystemID AS SOLAR_SYSTEM_ID, SOLAR_SYSTEMS.solarSystemName AS SOLAR_SYSTEM_NAME, SOLAR_SYSTEMS.security AS SOLAR_SYSTEM_SECURITY, REGIONS.regionID AS REGION_ID, REGIONS.regionName AS REGION_NAME, "
                    SQL = SQL & "FACILITY_TYPE_ID, typeName AS FACILITY_TYPE, RAM_ACTIVITIES.activityID AS ACTIVITY_ID, RAM_ACTIVITIES.activityName AS ACTIVITY_NAME, "
                    SQL = SQL & "FACILITY_TAX, baseMaterialMultiplier AS BASE_MM, baseTimeMultiplier AS BASE_TM, baseCostMultiplier AS BASE_CM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.materialMultiplier AS ADDITIONAL_MM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.timeMultiplier AS ADDITIONAL_TM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.costMultiplier AS ADDITIONAL_CM, "
                    SQL = SQL & "INVENTORY_GROUPS.groupID AS GROUP_ID, INVENTORY_GROUPS.groupName AS GROUP_NAME, 0 AS CATEGORY_ID, NULL AS CATEGORY_NAME, COST_INDEX, 0 AS OUTPOST "
                    SQL = SQL & "FROM INDUSTRY_FACILITIES, INVENTORY_TYPES, RAM_ASSEMBLY_LINE_STATIONS, RAM_ASSEMBLY_LINE_TYPES, RAM_ACTIVITIES, REGIONS, SOLAR_SYSTEMS, RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP, "
                    SQL = SQL & "INVENTORY_GROUPS, INDUSTRY_SYSTEMS_COST_INDICIES "
                    SQL = SQL & "WHERE INDUSTRY_FACILITIES.FACILITY_TYPE_ID = INVENTORY_TYPES.typeID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.REGION_ID = REGIONS.regionID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.SOLAR_SYSTEM_ID = SOLAR_SYSTEMS.solarSystemID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.FACILITY_ID = RAM_ASSEMBLY_LINE_STATIONS.stationID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPES.assemblyLineTypeID = RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.assemblyLineTypeID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPES.activityID = RAM_ACTIVITIES.activityID "
                    SQL = SQL & "AND INDUSTRY_SYSTEMS_COST_INDICIES.SOLAR_SYSTEM_ID = INDUSTRY_FACILITIES.SOLAR_SYSTEM_ID "
                    SQL = SQL & "AND INDUSTRY_SYSTEMS_COST_INDICIES.ACTIVITY_ID = RAM_ASSEMBLY_LINE_TYPES.activityID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_STATIONS.assemblyLineTypeID = RAM_ASSEMBLY_LINE_TYPES.assemblyLineTypeID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.groupID = INVENTORY_GROUPS.groupID "
                    SQL = SQL & "UNION "
                    SQL = SQL & "SELECT FACILITY_ID, FACILITY_NAME, "
                    SQL = SQL & "SOLAR_SYSTEMS.solarSystemID AS SOLAR_SYSTEM_ID, SOLAR_SYSTEMS.solarSystemName AS SOLAR_SYSTEM_NAME, SOLAR_SYSTEMS.security AS SOLAR_SYSTEM_SECURITY, REGIONS.regionID AS REGION_ID, REGIONS.regionName AS REGION_NAME, "
                    SQL = SQL & "FACILITY_TYPE_ID, typeName AS FACILITY_TYPE, RAM_ACTIVITIES.activityID AS ACTIVITY_ID, RAM_ACTIVITIES.activityName AS ACTIVITY_NAME, "
                    SQL = SQL & "FACILITY_TAX, baseMaterialMultiplier AS BASE_MM, baseTimeMultiplier AS BASE_TM, baseCostMultiplier AS BASE_CM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.materialMultiplier AS ADDITIONAL_MM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.timeMultiplier AS ADDITIONAL_TM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.costMultiplier AS ADDITIONAL_CM, "
                    SQL = SQL & "0 as GROUP_ID, NULL as GROUP_NAME, INVENTORY_CATEGORIES.categoryID AS CATEGORY_ID, INVENTORY_CATEGORIES.categoryName AS CATEGORY_NAME, COST_INDEX, 0 AS OUTPOST "
                    SQL = SQL & "FROM INDUSTRY_FACILITIES, INVENTORY_TYPES, RAM_ASSEMBLY_LINE_STATIONS, RAM_ASSEMBLY_LINE_TYPES, RAM_ACTIVITIES, REGIONS, SOLAR_SYSTEMS, RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY, "
                    SQL = SQL & "INVENTORY_CATEGORIES, INDUSTRY_SYSTEMS_COST_INDICIES "
                    SQL = SQL & "WHERE INDUSTRY_FACILITIES.FACILITY_TYPE_ID = INVENTORY_TYPES.typeID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.REGION_ID = REGIONS.regionID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.SOLAR_SYSTEM_ID = SOLAR_SYSTEMS.solarSystemID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.FACILITY_ID = RAM_ASSEMBLY_LINE_STATIONS.stationID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPES.assemblyLineTypeID = RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.assemblyLineTypeID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPES.activityID = RAM_ACTIVITIES.activityID "
                    SQL = SQL & "AND INDUSTRY_SYSTEMS_COST_INDICIES.SOLAR_SYSTEM_ID = INDUSTRY_FACILITIES.SOLAR_SYSTEM_ID "
                    SQL = SQL & "AND INDUSTRY_SYSTEMS_COST_INDICIES.ACTIVITY_ID = RAM_ASSEMBLY_LINE_TYPES.activityID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_STATIONS.assemblyLineTypeID = RAM_ASSEMBLY_LINE_TYPES.assemblyLineTypeID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.categoryID = INVENTORY_CATEGORIES.categoryID "
                    SQL = SQL & ") "
                    SQL = SQL & "WHERE (GROUP_NAME IS NOT NULL OR CATEGORY_NAME IS NOT NULL) "
                    SQL = SQL & "UNION "
                    SQL = SQL & "SELECT * FROM ( "
                    SQL = SQL & "SELECT FACILITY_ID, FACILITY_NAME, "
                    SQL = SQL & "SOLAR_SYSTEMS.solarSystemID AS SOLAR_SYSTEM_ID, SOLAR_SYSTEMS.solarSystemName AS SOLAR_SYSTEM_NAME, SOLAR_SYSTEMS.security AS SOLAR_SYSTEM_SECURITY, REGIONS.regionID AS REGION_ID, REGIONS.regionName AS REGION_NAME, "
                    SQL = SQL & "FACILITY_TYPE_ID, typeName AS FACILITY_TYPE, RAM_ACTIVITIES.activityID AS ACTIVITY_ID, RAM_ACTIVITIES.activityName AS ACTIVITY_NAME, "
                    SQL = SQL & "FACILITY_TAX, baseMaterialMultiplier AS BASE_MM, baseTimeMultiplier AS BASE_TM, baseCostMultiplier AS BASE_CM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.materialMultiplier AS ADDITIONAL_MM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.timeMultiplier AS ADDITIONAL_TM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.costMultiplier AS ADDITIONAL_CM, "
                    SQL = SQL & "INVENTORY_GROUPS.groupID AS GROUP_ID, INVENTORY_GROUPS.groupName AS GROUP_NAME, 0 AS CATEGORY_ID, NULL AS CATEGORY_NAME, COST_INDEX, 1 AS OUTPOST "
                    SQL = SQL & "FROM INDUSTRY_FACILITIES, INVENTORY_TYPES, RAM_INSTALLATION_TYPE_CONTENTS, RAM_ASSEMBLY_LINE_TYPES, RAM_ACTIVITIES, REGIONS, SOLAR_SYSTEMS, RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP, "
                    SQL = SQL & "INVENTORY_GROUPS, INDUSTRY_SYSTEMS_COST_INDICIES "
                    SQL = SQL & "WHERE FACILITY_TYPE_ID IN (21642,21644,21645,21646) "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.FACILITY_TYPE_ID = INVENTORY_TYPES.typeID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.FACILITY_TYPE_ID = RAM_INSTALLATION_TYPE_CONTENTS.installationTypeID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.REGION_ID = REGIONS.regionID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.SOLAR_SYSTEM_ID = SOLAR_SYSTEMS.solarSystemID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPES.assemblyLineTypeID = RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.assemblyLineTypeID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPES.activityID = RAM_ACTIVITIES.activityID "
                    SQL = SQL & "AND INDUSTRY_SYSTEMS_COST_INDICIES.SOLAR_SYSTEM_ID = INDUSTRY_FACILITIES.SOLAR_SYSTEM_ID "
                    SQL = SQL & "AND INDUSTRY_SYSTEMS_COST_INDICIES.ACTIVITY_ID = RAM_ASSEMBLY_LINE_TYPES.activityID "
                    SQL = SQL & "AND RAM_INSTALLATION_TYPE_CONTENTS.assemblyLineTypeID = RAM_ASSEMBLY_LINE_TYPES.assemblyLineTypeID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_GROUP.groupID = INVENTORY_GROUPS.groupID "
                    SQL = SQL & "UNION "
                    SQL = SQL & "SELECT FACILITY_ID, FACILITY_NAME, "
                    SQL = SQL & "SOLAR_SYSTEMS.solarSystemID AS SOLAR_SYSTEM_ID, SOLAR_SYSTEMS.solarSystemName AS SOLAR_SYSTEM_NAME, SOLAR_SYSTEMS.security AS SOLAR_SYSTEM_SECURITY, REGIONS.regionID AS REGION_ID, REGIONS.regionName AS REGION_NAME, "
                    SQL = SQL & "FACILITY_TYPE_ID, typeName AS FACILITY_TYPE, RAM_ACTIVITIES.activityID AS ACTIVITY_ID, RAM_ACTIVITIES.activityName AS ACTIVITY_NAME, "
                    SQL = SQL & "FACILITY_TAX, baseMaterialMultiplier AS BASE_MM, baseTimeMultiplier AS BASE_TM, baseCostMultiplier AS BASE_CM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.materialMultiplier AS ADDITIONAL_MM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.timeMultiplier AS ADDITIONAL_TM, "
                    SQL = SQL & "RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.costMultiplier AS ADDITIONAL_CM, "
                    SQL = SQL & "0 as GROUP_ID, NULL as GROUP_NAME, INVENTORY_CATEGORIES.categoryID AS CATEGORY_ID, INVENTORY_CATEGORIES.categoryName AS CATEGORY_NAME, COST_INDEX, 1 AS OUTPOST "
                    SQL = SQL & "FROM INDUSTRY_FACILITIES, INVENTORY_TYPES, RAM_INSTALLATION_TYPE_CONTENTS, RAM_ASSEMBLY_LINE_TYPES, RAM_ACTIVITIES, REGIONS, SOLAR_SYSTEMS, RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY, "
                    SQL = SQL & "INVENTORY_CATEGORIES, INDUSTRY_SYSTEMS_COST_INDICIES "
                    SQL = SQL & "WHERE FACILITY_TYPE_ID IN (21642,21644,21645,21646) "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.FACILITY_TYPE_ID = INVENTORY_TYPES.typeID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.FACILITY_TYPE_ID = RAM_INSTALLATION_TYPE_CONTENTS.installationTypeID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.REGION_ID = REGIONS.regionID "
                    SQL = SQL & "AND INDUSTRY_FACILITIES.SOLAR_SYSTEM_ID = SOLAR_SYSTEMS.solarSystemID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPES.assemblyLineTypeID = RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.assemblyLineTypeID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPES.activityID = RAM_ACTIVITIES.activityID "
                    SQL = SQL & "AND INDUSTRY_SYSTEMS_COST_INDICIES.SOLAR_SYSTEM_ID = INDUSTRY_FACILITIES.SOLAR_SYSTEM_ID "
                    SQL = SQL & "AND INDUSTRY_SYSTEMS_COST_INDICIES.ACTIVITY_ID = RAM_ASSEMBLY_LINE_TYPES.activityID "
                    SQL = SQL & "AND RAM_INSTALLATION_TYPE_CONTENTS.assemblyLineTypeID = RAM_ASSEMBLY_LINE_TYPES.assemblyLineTypeID "
                    SQL = SQL & "AND RAM_ASSEMBLY_LINE_TYPE_DETAIL_PER_CATEGORY.categoryID = INVENTORY_CATEGORIES.categoryID "
                    SQL = SQL & ") "
                    SQL = SQL & "WHERE (GROUP_NAME IS NOT NULL OR CATEGORY_NAME IS NOT NULL)"

                    Call ExecuteNonQuerySQL(SQL)

                    TempLabel.Text = "Indexing Stations Table..."
                    Application.DoEvents()

                    ' Index the table
                    SQL = "CREATE INDEX IDX_SF_OP_AID_CID_GID ON STATION_FACILITIES (OUTPOST, ACTIVITY_ID, CATEGORY_ID, GROUP_ID, REGION_NAME, SOLAR_SYSTEM_NAME)"
                    Call ExecuteNonQuerySQL(SQL)

                    SQL = "CREATE INDEX IDX_SF_OP_FN_AID_CID_GID ON STATION_FACILITIES (OUTPOST, FACILITY_NAME, ACTIVITY_ID, CATEGORY_ID, GROUP_ID, REGION_NAME, SOLAR_SYSTEM_NAME)"
                    Call ExecuteNonQuerySQL(SQL)

                    SQL = "CREATE INDEX IDX_SF_GID ON STATION_FACILITIES (GROUP_ID)"
                    Call ExecuteNonQuerySQL(SQL)

                    SQL = "CREATE INDEX IDX_SF_CID ON STATION_FACILITIES (CATEGORY_ID)"
                    Call ExecuteNonQuerySQL(SQL)

                    SQL = "CREATE INDEX IDX_SF_FN ON STATION_FACILITIES (FACILITY_NAME)"
                    Call ExecuteNonQuerySQL(SQL)

                    ' Finally, make the facility names table for easy look ups
                    ' First drop the table if it exists
                    SQL = "DROP TABLE IF EXISTS STATIONS"
                    Call ExecuteNonQuerySQL(SQL)

                    ' Create it
                    SQL = "CREATE TABLE STATIONS AS SELECT DISTINCT FACILITY_ID AS STATION_ID, FACILITY_NAME AS STATION_NAME, FACILITY_TYPE_ID AS STATION_TYPE_ID, "
                    SQL = SQL & "SOLAR_SYSTEM_ID, SOLAR_SYSTEM_SECURITY, REGION_ID FROM STATION_FACILITIES "
                    Call ExecuteNonQuerySQL(SQL)

                    ' Index the ID
                    SQL = "CREATE INDEX IDX_FN_FID ON STATIONS (STATION_ID)"
                    Call ExecuteNonQuerySQL(SQL)

                    ' Set the Cache Date to now plus the length since it's not sent in the file
                    Call SetCRESTCacheDate(IndustryFacilitiesField, CacheDate)

                    Call CommitSQLiteTransaction()

                End If
            End If
        End If

    End Sub

    ' For Industry Facilities
    Private Class IndustryFacilities
        '{"totalCount_str": "1005", "items":[], "pageCount": 1, "pageCount_str": "1", "totalCount": 1005}
        <JsonProperty("totalCount_str")> Public totalCount_str As String
        <JsonProperty("items")> Public items() As IndustryFacility
        <JsonProperty("pageCount")> Public pageCount As Integer
        <JsonProperty("pageCount_str")> Public pageCount_str As String
        <JsonProperty("totalCount")> Public totalCount As Integer
    End Class

    ' For Industry Facilities
    Private Class IndustryFacility
        '{"facilityID": 60012160, "solarSystem": {}, "name": "xxx", "region": {}, "tax": 0.1, "facilityID_str": "60012160", "owner": {}, "type": {} }, 
        <JsonProperty("facilityID")> Public facilityID As Long
        <JsonProperty("solarSystem")> Public solarSystem As IndustryFacilitySolarSystem
        <JsonProperty("name")> Public name As String
        <JsonProperty("region")> Public region As IndustryFacilityRegion
        <JsonProperty("tax")> Public tax As Double
        <JsonProperty("facilityID_str")> Public facilityID_str As String
        <JsonProperty("owner")> Public owner As IndustryFacilitiesOwner
        <JsonProperty("type")> Public type As IndustryActivities
    End Class

    ' For Industry Facilities
    Private Class IndustryFacilitiesOwner
        ' "owner": {"id": 1000123, "id_str": "1000123"},  
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("id")> Public id As Integer
    End Class

    ' For Industry Facilities
    Private Class IndustryActivities
        ' "type": {"id": 2500, "id_str": "2500"}
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("id")> Public id As Integer
    End Class

    Private Class IndustryFacilitySolarSystem
        ' "solarSystem": {"id": 30000049, "id_str": "30000049"}, 
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("id")> Public id As Long
    End Class

    ' For CREST Regions
    Private Class IndustryFacilityRegion
        ' "region": {"id": 10000001, "id_str": "10000001"}, 
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("id")> Public id As Long
    End Class

    ' /industry/systems/ (cache: 1 hour)
    ' vnd.ccp.eve.IndustrySystemCollection-v1
    ' Lists the cost index for installing industry jobs per type of activity. This does not include wormhole space.

    ' Gets the CREST file from CCP for current Cost Indexes of Industry Systems and updates the EVEIPH DB with the values
    Public Sub UpdateIndustrySystemsCostIndex(Optional ByRef UpdateLabel As Label = Nothing, Optional ByRef PB As ProgressBar = Nothing)
        Dim IndustrySystemsIndex As IndustrySystemCostIndicies
        Dim SQL As String
        Dim CacheDate As Date

        Dim SolarSystemID As Long
        Dim SolarSystemName As String

        Dim TempLabel As Label
        Dim TempPB As ProgressBar

        If IsNothing(UpdateLabel) Then
            TempLabel = New Label
        Else
            TempLabel = UpdateLabel
        End If

        If IsNothing(PB) Then
            TempPB = New ProgressBar
        Else
            TempPB = PB
        End If

        ' First look up the cache date to see if it's time to run the update
        CacheDate = GetCRESTCacheDate(IndustrySystemsField)

        ' If it's later than now, update
        If CacheDate <= Now Then

            TempLabel.Text = "Downloading System Index Data..."
            Application.DoEvents()

            ' Dump the file into the Specializations object
            IndustrySystemsIndex = JsonConvert.DeserializeObject(Of IndustrySystemCostIndicies) _
                (GetJSONFile(CRESTRootServerURL & CRESTIndustrySystems, CacheDate))

            ' Read in the data
            If Not IsNothing(IndustrySystemsIndex) Then
                If IndustrySystemsIndex.items.Count > 0 Then
                    Call BeginSQLiteTransaction()

                    TempLabel.Text = "Saving System Index Data..."
                    TempPB.Minimum = 0
                    TempPB.Value = 0
                    TempPB.Maximum = IndustrySystemsIndex.totalCount - 1
                    TempPB.Visible = True
                    Application.DoEvents()

                    ' Delete the old records first
                    Call ExecuteNonQuerySQL("DELETE FROM INDUSTRY_SYSTEMS_COST_INDICIES")

                    ' Now read through all the output items and input them into the DB
                    For i = 0 To IndustrySystemsIndex.totalCount - 1

                        SolarSystemID = IndustrySystemsIndex.items(i).solarSystem.id
                        SolarSystemName = IndustrySystemsIndex.items(i).solarSystem.name

                        For j = 0 To IndustrySystemsIndex.items(i).systemCostIndices.Count - 1
                            With IndustrySystemsIndex.items(i).systemCostIndices(j)
                                ' Insert the record
                                SQL = "INSERT INTO INDUSTRY_SYSTEMS_COST_INDICIES VALUES(" & CStr(SolarSystemID) & ",'" & SolarSystemName & "',"
                                SQL = SQL & CStr(.activityID) & ",'" & .activityName & "'," & CStr(.costIndex) & ")"
                                Call ExecuteNonQuerySQL(SQL)
                            End With
                        Next

                        TempPB.Value = i
                        Application.DoEvents()
                    Next

                    TempPB.Visible = False

                    ' Rebuild indexes
                    Call ExecuteNonQuerySQL("REINDEX IDX_ISCI_AID_SSID")

                    ' Set the Cache Date to now plus the length since it's not sent in the file
                    Call SetCRESTCacheDate(IndustrySystemsField, CacheDate)
                    ' Done updating
                    Call CommitSQLiteTransaction()
                End If
            End If
        End If

    End Sub

    ' For Industry System Cost Indicies
    Private Class IndustrySystemCostIndicies
        '{"totalCount_str": "1005", "items":[], "pageCount": 1, "pageCount_str": "1", "totalCount": 1005}
        <JsonProperty("totalCount_str")> Public totalCount_str As String
        <JsonProperty("items")> Public items() As systemCostIndices
        <JsonProperty("pageCount")> Public pageCount As Integer
        <JsonProperty("pageCount_str")> Public pageCount_str As String
        <JsonProperty("totalCount")> Public totalCount As Integer
    End Class

    ' For Industry System Cost Indicies
    Private Class systemCostIndices
        '{"systemCostIndices": [], "solarSystem": {}}
        <JsonProperty("systemCostIndices")> Public systemCostIndices() As costIndex
        <JsonProperty("solarSystem")> Public solarSystem As SolarSystem
    End Class

    ' For Industry System Cost Indicies
    Private Class costIndex
        ''{"costIndex": 0.02832645090482675, "activityID": 1, "activityID_str": "1", "activityName": "Manufacturing"}
        <JsonProperty("costIndex")> Public costIndex As Double
        <JsonProperty("activityID")> Public activityID As Integer
        <JsonProperty("activityID_str")> Public activityID_str As String
        <JsonProperty("activityName")> Public activityName As String
    End Class

    ' /market/prices/ (cache: 23 hours)
    ' vnd.ccp.eve.MarketTypePriceCollection-v1
    ' Returns the list of trade-able types and their average market price, as shown in the inventory UI in the EVE client. Also includes an adjusted market price which is used in industry calculations.

    ' Gets the CREST file from CCP for current Market Prices and updates the EVEIPH DB with the values
    Public Sub UpdateMarketPrices(Optional ByRef UpdateLabel As Label = Nothing, Optional ByRef PB As ProgressBar = Nothing)
        Dim MarketPricesOutput As MarketPrices
        Dim SQL As String
        Dim CacheDate As Date
        Dim TempLabel As Label
        Dim TempPB As ProgressBar

        If IsNothing(UpdateLabel) Then
            TempLabel = New Label
        Else
            TempLabel = UpdateLabel
        End If

        If IsNothing(PB) Then
            TempPB = New ProgressBar
        Else
            TempPB = PB
        End If

        ' First look up the cache date to see if it's time to run the update
        CacheDate = GetCRESTCacheDate(MarketPricesField)

        ' If it's later than now, update
        If CacheDate <= Now Then

            TempLabel.Text = "Downloading Market Price Data..."
            Application.DoEvents()

            ' Dump the file into the Specializations object
            MarketPricesOutput = JsonConvert.DeserializeObject(Of MarketPrices) _
                (GetJSONFile(CRESTRootServerURL & CRESTMarketPrices, CacheDate))

            ' Read in the data
            If Not IsNothing(MarketPricesOutput) Then
                If MarketPricesOutput.items.Count > 0 Then
                    Call BeginSQLiteTransaction()

                    ' Clear the old records first
                    Call ExecuteNonQuerySQL("UPDATE ITEM_PRICES SET ADJUSTED_PRICE = 0, AVERAGE_PRICE = 0")

                    TempLabel.Text = "Saving Market Price Data..."
                    TempPB.Minimum = 0
                    TempPB.Value = 0
                    TempPB.Maximum = MarketPricesOutput.totalCount - 1
                    TempPB.Visible = True
                    Application.DoEvents()

                    ' Now read through all the output items and update them in ITEM_PRICES
                    For i = 0 To MarketPricesOutput.totalCount - 1
                        With MarketPricesOutput.items(i)
                            SQL = "UPDATE ITEM_PRICES SET ADJUSTED_PRICE = " & CStr(.adjustedPrice) & ", AVERAGE_PRICE = " & CStr(.averagePrice)
                            SQL = SQL & " WHERE ITEM_ID = " & CStr(.type.id)
                            Call ExecuteNonQuerySQL(SQL)
                        End With

                        TempPB.Value = i
                        Application.DoEvents()

                    Next

                    ' Set the Cache Date to now plus the length since it's not sent in the file
                    Call SetCRESTCacheDate(MarketPricesField, CacheDate)
                    ' Done updating
                    Call CommitSQLiteTransaction()
                End If
            End If
        End If

    End Sub

    ' For Market Prices
    Private Class MarketPrices
        '{"totalCount_str": "1005", "items":[], "pageCount": 1, "pageCount_str": "1", "totalCount": 1005}
        <JsonProperty("totalCount_str")> Public totalCount_str As String
        <JsonProperty("items")> Public items() As MarketAdjustedPrice
        <JsonProperty("pageCount")> Public pageCount As Integer
        <JsonProperty("pageCount_str")> Public pageCount_str As String
        <JsonProperty("totalCount")> Public totalCount As Integer
    End Class

    ' For Market Prices
    Private Class MarketAdjustedPrice
        '{"adjustedPrice": 567464.0783, "averagePrice": 565523.7836, "type": {} }
        <JsonProperty("adjustedPrice")> Public adjustedPrice As Double
        <JsonProperty("averagePrice")> Public averagePrice As Double
        <JsonProperty("type")> Public type As MarketPriceType
    End Class

    ' For Market Prices
    Private Class MarketPriceType
        '"type": {"id_str": "32772", "href": "http://public-crest-sisi.testeveonline.com/types/32772/", "id": 32772, "name": "Medium Ancillary Shield Booster"}
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("href")> Public href As String
        <JsonProperty("id")> Public id As Long
        <JsonProperty("name")> Public name As String
    End Class

    ' Stuff for all functions

    ' For CREST Solar Systems
    Private Class SolarSystem
        '{"id_str": "30001743", "href": "http://public-crest-sisi.testeveonline.com/solarsystems/30001743/", "id": 30001743, "name": "JUE-DX"}, 
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("href")> Public href As String
        <JsonProperty("id")> Public id As Long
        <JsonProperty("name")> Public name As String
    End Class

    ' For CREST Regions
    Private Class Region
        '{"id_str": "30001743", "href": "http://public-crest-sisi.testeveonline.com/solarsystems/30001743/", "id": 30001743, "name": "The Forge"}, 
        <JsonProperty("id_str")> Public id_str As String
        <JsonProperty("href")> Public href As String
        <JsonProperty("id")> Public id As Long
        <JsonProperty("name")> Public name As String
    End Class

    ' For all CREST where key is just a web link
    Private Class hrefKey
        '"capsuleer": {"href": "http://public-crest-sisi.testeveonline.com/characters/1047420507/capsuleer/"}
        <JsonProperty("href")> Public href As String
    End Class

    ' Downloads the JSON file sent and saves it to the location, then imports it into a string to return
    Private Function GetJSONFile(ByVal URL As String, ByRef CacheDate As Date) As String
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader
        Dim MaxAge As String
        Dim CallStatus As String

        Try
            ' Create the web request  
            request = DirectCast(WebRequest.Create(URL), HttpWebRequest)

            ' Get response  
            response = DirectCast(request.GetResponse(), HttpWebResponse)

            ' Use later to see what calls hit or fail
            CallStatus = response.Headers.Get("X-Cache-Status")

            ' Get the max-age from cache control
            MaxAge = response.Headers.Get("Cache-Control")
            ' Parse max age for seconds
            Dim Seconds As Integer = CInt(MaxAge.Substring(MaxAge.IndexOf("=") + 1))

            ' Get the response stream into a reader  
            reader = New StreamReader(response.GetResponseStream())

            ' Set the cache date by ref
            CacheDate = DateAdd(DateInterval.Second, Seconds, CDate(response.Headers.Get("Date")))

            Return reader.ReadToEnd

        Finally
            If Not response Is Nothing Then response.Close()
        End Try

    End Function

    ' Gets the CREST Cache Date for the field name sent in the EVEIPH_DATA table
    Private Function GetCRESTCacheDate(UpdateField As String) As Date
        Dim SQL As String
        Dim readerData As SQLite.SQLiteDataReader
        Dim RefreshDate As Date

        SQL = "SELECT " & UpdateField & " FROM EVEIPH_DATA"
        DBCommand = New SQLiteCommand(SQL, DB)
        readerData = DBCommand.ExecuteReader

        If readerData.Read Then
            If Not IsDBNull(readerData.GetValue(0)) Then
                If readerData.GetString(0) = "" Then
                    RefreshDate = NoDate
                Else
                    RefreshDate = CDate(readerData.GetString(0))
                End If
            Else
                RefreshDate = NoDate
            End If
        Else
            RefreshDate = NoDate
        End If

        readerData.Close()
        readerData = Nothing

        Return RefreshDate

    End Function

    ' Sets the CREST Cache Date for the field name sent in the EVEIPH_DATA table
    Private Sub SetCRESTCacheDate(UpdateField As String, CacheDate As Date)
        Dim SQL As String
        Dim readerData As SQLite.SQLiteDataReader

        ' Update the cache for this CREST file
        SQL = "SELECT " & UpdateField & " FROM EVEIPH_DATA"
        DBCommand = New SQLiteCommand(SQL, DB)
        readerData = DBCommand.ExecuteReader

        If readerData.Read Then
            SQL = "UPDATE EVEIPH_DATA SET " & UpdateField & " = '" & Format(CacheDate, SQLiteDateFormat) & "'"
            ExecuteNonQuerySQL(SQL)
        Else
            SQL = "INSERT INTO EVEIPH_DATA (" & UpdateField & ") VALUES ('" & Format(CacheDate, SQLiteDateFormat) & "')"
            ExecuteNonQuerySQL(SQL)
        End If
    End Sub

End Class

