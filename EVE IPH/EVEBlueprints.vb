
Imports System.Data.SQLite

Public Class EVEBlueprints
    Private BlueprintList As List(Of EVEBlueprint)
    Private KeyData As APIKeyData
    Private CorpID As Long

    Private CacheDate As Date

    Public Sub New(Optional ByVal Key As APIKeyData = Nothing, Optional ByVal CorporationID As Long = 0)

        If IsNothing(Key) Then
            KeyData = New APIKeyData
        Else
            KeyData = Key
        End If

        BlueprintList = New List(Of EVEBlueprint)

        ' Set for corp Blueprints
        CorpID = CorporationID

        ' Default to this until we set it
        CacheDate = NoExpiry

    End Sub

    ' Loads all blueprints for the character from the DB
    Public Sub LoadBlueprints(BlueprintType As ScanType, UpdatefromAPI As Boolean)
        Dim SQL As String
        Dim readerBlueprints As SQLiteDataReader
        Dim TempBlueprint As EVEBlueprint
        Dim Blueprints As New List(Of EVEBlueprint)
        Dim SearchID As Long

        If Not KeyData.Access Then
            ' They don't have access to the api so leave
            Exit Sub
        End If

        ' Update Industry Blueprints first
        Call UpdateBlueprints(BlueprintType, UpdatefromAPI)

        If BlueprintType = ScanType.Personal Then
            SearchID = KeyData.ID
        Else
            SearchID = CorpID
        End If

        ' Load the blueprints
        SQL = "SELECT ITEM_ID, LOCATION_ID, BLUEPRINT_ID, BLUEPRINT_NAME, FLAG_ID, QUANTITY, ME, TE, "
        SQL = SQL & "RUNS, BP_TYPE, OWNED, SCANNED, FAVORITE, ADDITIONAL_COSTS "
        SQL = SQL & "FROM OWNED_BLUEPRINTS WHERE USER_ID = " & SearchID

        DBCommand = New SQLiteCommand(SQL, DB)
        readerBlueprints = DBCommand.ExecuteReader

        While readerBlueprints.Read

            If IsDBNull(readerBlueprints.GetValue(0)) Then
                TempBlueprint.itemID = 0
            Else
                TempBlueprint.itemID = readerBlueprints.GetInt64(0)
            End If

            TempBlueprint.locationID = readerBlueprints.GetInt64(1)
            TempBlueprint.typeID = readerBlueprints.GetInt64(2)
            TempBlueprint.typeName = readerBlueprints.GetString(3)
            TempBlueprint.flagID = readerBlueprints.GetInt32(4)
            TempBlueprint.quantity = readerBlueprints.GetInt32(5)
            TempBlueprint.timeEfficiency = readerBlueprints.GetDouble(6)
            TempBlueprint.materialEfficiency = readerBlueprints.GetDouble(7)
            TempBlueprint.runs = readerBlueprints.GetInt32(8)
            TempBlueprint.BPType = readerBlueprints.GetInt32(9)
            TempBlueprint.Owned = CBool(readerBlueprints.GetInt32(10))
            TempBlueprint.Scanned = CBool(readerBlueprints.GetInt32(11))
            TempBlueprint.Favorite = CBool(readerBlueprints.GetInt32(12))
            TempBlueprint.AdditionalCosts = readerBlueprints.GetDouble(13)

            ' Insert blueprint
            Blueprints.Add(TempBlueprint)

        End While

        readerBlueprints.Close()
        DBCommand = Nothing
        readerBlueprints = Nothing

        BlueprintList = Blueprints

    End Sub

    ' Updates Blueprints from API for the character/corp and inserts them into the Database for later queries
    Private Sub UpdateBlueprints(BlueprintType As ScanType, UpdateAPI As Boolean)
        Dim readerBlueprints As SQLiteDataReader
        Dim SQL As String
        Dim RefreshDate As Date ' To check the update of the API.
        Dim API As New EVEAPI
        Dim IndyBlueprints As New List(Of EVEBlueprint)
        Dim SearchID As Long
        Dim InsertBP As Boolean
        Dim IgnoreBP As Boolean
        Dim ScannedFlag As Integer

        ' See if we are doing an API update 
        If Not UpdateAPI Then
            Exit Sub
        End If

        ' Look up the industry Blueprints cache date first, if past the date, update the database
        SQL = "SELECT BLUEPRINTS_CACHED_UNTIL FROM API "
        If BlueprintType = ScanType.Personal Then
            SQL = SQL & "WHERE CHARACTER_ID = " & KeyData.ID
            SQL = SQL & " AND API_TYPE NOT IN ('Corporation', 'Old Key')"
            SearchID = KeyData.ID
            ScannedFlag = 1
        Else
            SQL = SQL & "WHERE CORPORATION_ID = " & CorpID
            SQL = SQL & " AND API_TYPE = 'Corporation'"
            SearchID = CorpID
            ScannedFlag = 2
        End If

        DBCommand = New SQLiteCommand(SQL, DB)
        readerBlueprints = DBCommand.ExecuteReader

        If readerBlueprints.Read Then
            If Not IsDBNull(readerBlueprints.GetValue(0)) Then
                If readerBlueprints.GetString(0) = "" Then
                    RefreshDate = NoDate
                Else
                    RefreshDate = CDate(readerBlueprints.GetString(0))
                End If
            Else
                RefreshDate = NoDate
            End If
        Else
            RefreshDate = NoDate
        End If

        ' See if we refresh the data 
        If RefreshDate <= DateTime.UtcNow Then

            IndyBlueprints = API.GetBlueprints(KeyData, BlueprintType, CacheDate)

            If Not NoAPIError(API.GetErrorText, "Character") Then
                ' Errored, exit
                Exit Sub
            End If

            ' Begin session
            Call BeginSQLiteTransaction()

            ' Update the cache date
            SQL = "UPDATE API SET BLUEPRINTS_CACHED_UNTIL = '" & Format(CacheDate, SQLiteDateFormat)

            If BlueprintType = ScanType.Personal Then
                SQL = SQL & "' WHERE CHARACTER_ID = " & SearchID
                SQL = SQL & " AND API_TYPE NOT IN ('Corporation', 'Old Key')"
            Else
                SQL = SQL & "' WHERE CORPORATION_ID = " & SearchID
                SQL = SQL & " AND API_TYPE = 'Corporation'"
            End If

            Call ExecuteNonQuerySQL(SQL)

            If Not IsNothing(IndyBlueprints) Then

                ' Insert blueprint data
                For i = 0 To IndyBlueprints.Count - 1
                    ' First make sure it's not already in there
                    With IndyBlueprints(i)

                        ' For now, only include unique BPs until I get the multiple BP support done - use Max ME for the determination or Max TE if they are the same ME
                        SQL = "SELECT ME, TE FROM OWNED_BLUEPRINTS WHERE BLUEPRINT_ID = " & .typeID & " AND USER_ID = " & SearchID
                        DBCommand = New SQLiteCommand(SQL, DB)
                        readerBlueprints = DBCommand.ExecuteReader
                        readerBlueprints.Read()

                        If readerBlueprints.HasRows Then
                            ' If greater ME, or the ME is equal and TE is greater, update it
                            If readerBlueprints.GetDouble(0) > IndyBlueprints(i).materialEfficiency Then
                                InsertBP = False
                                IgnoreBP = False
                            ElseIf readerBlueprints.GetDouble(0) = IndyBlueprints(i).materialEfficiency And readerBlueprints.GetDouble(1) > IndyBlueprints(i).timeEfficiency Then
                                InsertBP = False
                                IgnoreBP = False
                            Else
                                ' We don't want to do anything with this bp
                                IgnoreBP = True
                            End If
                        Else
                            IgnoreBP = False
                            InsertBP = True
                        End If

                        If Not IgnoreBP Then
                            If InsertBP Then
                                ' Full insert always mark owned unless it's a BPC
                                Dim OwnedFlag As String

                                If .BPType = BPType.Copy Then
                                    OwnedFlag = "0"
                                Else
                                    OwnedFlag = "1"
                                End If
                                SQL = "INSERT INTO OWNED_BLUEPRINTS (USER_ID, ITEM_ID, LOCATION_ID, BLUEPRINT_ID, BLUEPRINT_NAME, FLAG_ID, "
                                SQL = SQL & "QUANTITY, ME, TE, RUNS, BP_TYPE, OWNED, SCANNED, FAVORITE, ADDITIONAL_COSTS) "
                                SQL = SQL & "VALUES (" & SearchID & "," & .itemID & "," & .locationID & "," & .typeID & ",'" & FormatDBString(.typeName) & "',"
                                SQL = SQL & .flagID & "," & .quantity & "," & .materialEfficiency & "," & .timeEfficiency & "," & .runs & "," & .BPType & "," & OwnedFlag & "," & CStr(ScannedFlag) & ", 0, 0)"
                            Else
                                ' Update the BP 
                                SQL = "UPDATE OWNED_BLUEPRINTS SET "
                                SQL = SQL & "LOCATION_ID = " & .locationID & ","
                                SQL = SQL & "FLAG_ID = " & .flagID & ","
                                SQL = SQL & "ME = " & .materialEfficiency & ","
                                SQL = SQL & "TE = " & .timeEfficiency & ","
                                SQL = SQL & "RUNS = " & .runs & ","
                                SQL = SQL & "BLUEPRINT_NAME = '" & FormatDBString(.typeName) & "' " ' If it changes
                                ' Search with ITEM_ID
                                SQL = SQL & "WHERE ITEM_ID = " & .itemID & " AND USER_ID = " & SearchID

                            End If

                            Call ExecuteNonQuerySQL(SQL)
                        End If

                        readerBlueprints.Close()
                        readerBlueprints = Nothing

                    End With
                Next

                DBCommand = Nothing

            End If

            Call CommitSQLiteTransaction()

        End If

    End Sub

    ReadOnly Property CachedUntil() As Date
        Get
            Return CacheDate
        End Get
    End Property

End Class

' For outputing lists of blueprints
Public Structure EVEBlueprint
    Dim itemID As Long
    Dim locationID As Long
    Dim typeID As Long
    Dim typeName As String
    Dim flagID As Integer
    Dim quantity As Integer
    Dim timeEfficiency As Double
    Dim materialEfficiency As Double
    Dim runs As Integer
    Dim BPType As Integer
    Dim Owned As Boolean
    Dim Scanned As Boolean
    Dim Favorite As Boolean
    Dim AdditionalCosts As Double
End Structure
