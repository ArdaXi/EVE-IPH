
' Stores a list of materials and associated functions
Public Class Materials
    Implements ICloneable

    ' The List of Materials
    Private MaterialList() As Material

    ' Total Cost of materials in the list
    Private TotalMaterialsCost As Double
    ' Total Volume of materials in the list
    Private TotalMaterialsVolume As Double

    ' Constructor
    Public Sub New()
        TotalMaterialsCost = 0
        TotalMaterialsVolume = 0

        MaterialList = Nothing
    End Sub

    ' Destructor
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    ' For doing a deep copy of Materials
    Public Function Clone() As Object Implements ICloneable.Clone
        Dim CopyOfMe = New Materials
        Dim TempMat As Material

        If Not IsNothing(MaterialList) Then
            For i = 0 To MaterialList.Count - 1
                TempMat = CType(MaterialList(i).Clone(), Material)
                CopyOfMe.InsertMaterial(TempMat)
            Next

            CopyOfMe.TotalMaterialsCost = Me.TotalMaterialsCost
            CopyOfMe.TotalMaterialsVolume = Me.TotalMaterialsVolume
        Else
            CopyOfMe = Nothing
        End If

        Return CopyOfMe

    End Function

    ' Resets the List
    Public Sub Clear()
        TotalMaterialsCost = 0
        TotalMaterialsVolume = 0

        MaterialList = Nothing
    End Sub

    ' Searches the list and finds then returns a material for the name part sent
    Public Function SearchListbyName(ByVal SearchText As String, Optional ExactSearch As Boolean = False) As Material

        If Not IsNothing(MaterialList) Then
            For i = 0 To MaterialList.Count - 1
                If ExactSearch Then
                    If MaterialList(i).GetMaterialName = SearchText Then
                        Return MaterialList(i)
                    End If
                Else ' Look for partial string
                    If InStr(MaterialList(i).GetMaterialName, SearchText) <> 0 Then
                        Return MaterialList(i)
                    End If
                End If
            Next
        End If

        Return Nothing

    End Function

    ' Just adds a Full list to the Class
    Public Sub InsertMaterialList(ByVal SentList As Material())
        Dim i As Integer

        If Not IsNothing(SentList) Then
            For i = 0 To (SentList.Count - 1)
                Call InsertMaterial(SentList(i))
            Next
        End If

    End Sub

    ' Removes a full list of materials from the list
    Public Sub RemoveMaterialList(ByVal SentList As Material())
        Dim i As Integer

        If Not IsNothing(SentList) Then
            For i = 0 To (SentList.Count - 1)
                Call RemoveMaterial(SentList(i))
            Next
        End If

    End Sub

    ' Inserts a Single material into list
    Public Sub InsertMaterial(ByVal SentMaterial As Material)
        Dim i As Integer
        Dim TempMaterialArray() As Material
        Dim TempMat As Material

        Dim inList As Boolean
        Dim InListIndex As Integer

        inList = False

        ' Three possibilities
        ' 1 - No list created yet
        ' 2 - List created, not in list, add material
        ' 3 - List created, in list, need to update

        ' See if the material is in the list
        If Not IsNothing(MaterialList) Then
            For i = 0 To MaterialList.Count - 1
                ' Check name and group, the group is used by shopping list to store meta levels for now. ME values may be different as well that distinguishes the material
                If MaterialList(i).GetMaterialName = SentMaterial.GetMaterialName _
                    And MaterialList(i).GetMaterialGroup = SentMaterial.GetMaterialGroup _
                    And MaterialList(i).GetItemME = SentMaterial.GetItemME Then
                    InListIndex = i
                    inList = True
                    Exit For
                End If
            Next
        End If

        ' No List yet
        If IsNothing(MaterialList) Then
            ReDim MaterialList(0)
            ' Add material
        ElseIf inList Then ' In the list, just update quantity
            ' Just update quantity if in the List, material function will update volume and cost
            MaterialList(i).AddQuantity(SentMaterial.GetQuantity)
            ' Update this lists total cost with this new material
            TotalMaterialsCost = TotalMaterialsCost + SentMaterial.GetTotalCost
            TotalMaterialsVolume = TotalMaterialsVolume + SentMaterial.GetTotalVolume
            Exit Sub
        Else ' New record, Copy old array, rebuild new with new record on end
            ' Build the temp array
            ReDim TempMaterialArray(MaterialList.Count - 1)

            ' Copy old list
            TempMaterialArray = MaterialList

            ' Build the new array with one new record
            ReDim MaterialList(MaterialList.Count)

            ' Copy old data
            For i = 0 To MaterialList.Count - 2
                With TempMaterialArray(i)
                    TempMat = New Material(.GetMaterialTypeID, .GetMaterialName, .GetMaterialGroup, .GetQuantity, .GetVolume, .GetCostPerItem, .GetItemME, .GetBuildItem, .GetItemType)
                End With

                ' Set the new mat
                MaterialList(i) = TempMat
            Next

        End If

        ' Add the material at the end
        With SentMaterial
            TempMat = New Material(.GetMaterialTypeID, .GetMaterialName, .GetMaterialGroup, .GetQuantity, .GetVolume, .GetCostPerItem, .GetItemME, .GetBuildItem, .GetItemType)
            MaterialList(i) = TempMat
        End With

        ' Update the total cost of the class
        TotalMaterialsCost = TotalMaterialsCost + SentMaterial.GetTotalCost

        ' Update the total material volume for the list
        TotalMaterialsVolume = TotalMaterialsVolume + SentMaterial.GetTotalVolume

    End Sub

    ' Removes a Single material from the list
    Public Sub RemoveMaterial(ByVal SentMaterial As Material)
        Dim i As Integer
        Dim TempMaterialArray() As Material
        Dim TempMat As Material

        Dim InList As Boolean = False
        Dim InListIndex As Integer

        Dim PastRemovedItem As Boolean

        InList = False

        ' Make sure they send a valid material
        If IsNothing(SentMaterial) Then
            Exit Sub
        End If

        If IsNothing(MaterialList) Then
            ' There is no list yet
            Exit Sub
        Else
            ' See if the material is in the list
            For i = 0 To MaterialList.Count - 1
                If MaterialList(i).GetMaterialName = SentMaterial.GetMaterialName Then
                    InListIndex = i
                    InList = True
                    ' Found it, save location and exit loop
                    Exit For
                End If
            Next
        End If

        If InList Then ' In the list, check to see if the quantity is the same (name is), if so just remove, else update

            If MaterialList(i).GetQuantity = SentMaterial.GetQuantity Then
                ' Its in the list and all the quantity is the same, so remove

                ' Build the material array to save old mats
                ReDim TempMaterialArray(MaterialList.Count - 1)

                ' Copy old list
                TempMaterialArray = MaterialList

                ' Build the new array with one less record
                ReDim MaterialList(MaterialList.Count - 2)
                PastRemovedItem = False

                ' Copy old data until we get to the one we want to remove, then skip over it
                For i = 0 To TempMaterialArray.Count - 1
                    If TempMaterialArray(i).GetMaterialName <> SentMaterial.GetMaterialName Then
                        With TempMaterialArray(i)
                            TempMat = New Material(.GetMaterialTypeID, .GetMaterialName, .GetMaterialGroup, .GetQuantity, .GetVolume, .GetCostPerItem, .GetItemME, .GetBuildItem, .GetItemType)
                        End With

                        ' Set the new mat
                        If Not PastRemovedItem Then
                            MaterialList(i) = TempMat
                        Else
                            MaterialList(i - 1) = TempMat
                        End If
                    Else
                        PastRemovedItem = True
                    End If
                Next

            Else ' Update the quantity of the item in the list
                ' Just update quantity (negative to remove), material function will update volume and cost
                MaterialList(i).AddQuantity(-1 * SentMaterial.GetQuantity)
            End If

        End If

        ' Update the total cost of the class
        TotalMaterialsCost = TotalMaterialsCost - SentMaterial.GetTotalCost

        ' Update the total material volume for the list
        TotalMaterialsVolume = TotalMaterialsVolume - SentMaterial.GetTotalVolume

    End Sub

    ' Resets the value of the list to the sent value
    Public Sub ResetTotalValue(ByVal TotalValue As Double)
        TotalMaterialsCost = TotalValue
    End Sub

    ' Adds value to the total value of the list 
    Public Sub AddTotalValue(ByVal TotalValue As Double)
        TotalMaterialsCost = TotalMaterialsCost + TotalValue
    End Sub

    ' Adds volume to the total volume of the list
    Public Sub AddTotalVolume(ByVal TotalVolume As Double)
        TotalMaterialsVolume = TotalMaterialsVolume + TotalVolume
    End Sub

    ' "Adds" taxes to the total materials - i.e. takes off the taxes for selling these materials
    Public Sub AdjustTaxedPrice(ByVal TotalTax As Double)
        TotalMaterialsCost = TotalMaterialsCost - TotalTax
    End Sub

    ' Returns the list of Materials
    Public Function GetMaterialList() As Material()
        Return MaterialList
    End Function

    ' Sorts the Materials by quantity decending (Add more options later)
    Public Sub SortMaterialListByQuantity()
        If Not IsNothing(MaterialList) Then
            If MaterialList.Count - 1 > 0 Then
                ' Sort the list by quantity
                Call SortListDesc(MaterialList, 0, MaterialList.Count - 1)
            End If
        End If
    End Sub

    ' Returns the list in a clipboard format with CSV as an option
    Public Function GetClipboardList(ByVal ExportTextFormat As String, ByVal IgnorePriceVolume As Boolean, Optional IgnoreME As Boolean = False) As String
        Dim i As Integer
        Dim OutputString As String
        Dim MatName As String
        Dim DataInterfaces As String = ""
        Dim OutputME As String
        Dim DataInterfaceCost As Double
        Dim DataInterfaceVolume As Double
        Dim RelicDecryptorText As String = ""
        Dim Separator As String = ""

        If Not IsNothing(MaterialList) Then

            If ExportTextFormat = CSVDataExport Then
                OutputString = "Material, Quantity, ME, Decryptor/Relic, Cost Per Item, Total Cost" & vbCrLf
                Separator = ", "
            ElseIf ExportTextFormat = SSVDataExport Then
                OutputString = "Material; Quantity; ME; Decryptor/Relic; Cost Per Item; Total Cost" & vbCrLf
                Separator = "; "
            Else ' Default
                OutputString = "Material - Quantity" & vbCrLf
            End If

            ' Loop through all materials
            For i = 0 To MaterialList.Count - 1

                MatName = MaterialList(i).GetMaterialName

                ' Don't include data interfaces in the final output - this his poorly hacked but hacked nonetheless
                If MatName.Contains("Data Interface") Then
                    ' Add a crlf if not set
                    If DataInterfaces = "" Then
                        DataInterfaces = vbCrLf
                    End If
                    ' Add these to the end of the list if invention materials
                    DataInterfaces = DataInterfaces & "Uses " & MatName & vbCrLf
                    DataInterfaceCost = DataInterfaceCost + MaterialList(i).GetTotalCost
                    DataInterfaceVolume = DataInterfaceCost + MaterialList(i).GetTotalVolume
                    GoTo SkipFormat
                Else
                    DataInterfaceCost = 0
                    DataInterfaceVolume = 0
                End If

                If MaterialList(i).GetMaterialGroup.Contains("|") Then
                    ' We have a material from the shopping list, with three values in the material group
                    '.BuildType & "|" & .DecryptorRelic
                    ' Parse the fields
                    Dim ItemColumns As String() = MaterialList(i).GetMaterialGroup.Split(New [Char]() {"|"c})

                    If ItemColumns(1) <> None And ItemColumns(1) <> "" Then
                        RelicDecryptorText = ItemColumns(1)
                    Else
                        RelicDecryptorText = ""
                    End If
                End If

                If ExportTextFormat = CSVDataExport Or ExportTextFormat = SSVDataExport Then
                    ' Format output for no commas in prices or quantity
                    If IgnoreME Then
                        OutputME = "-"
                    Else
                        OutputME = MaterialList(i).GetItemME
                    End If

                    OutputString = OutputString & MatName & Separator & CStr(MaterialList(i).GetQuantity) & Separator & OutputME & Separator & RelicDecryptorText _
                        & Separator & CStr(MaterialList(i).GetCostPerItem) & Separator & CStr(MaterialList(i).GetTotalCost) & vbCrLf

                Else
                    OutputString = OutputString & MatName

                    If Not IgnoreME Then
                        If MaterialList(i).GetItemME <> "-" Then
                            OutputString = OutputString & " (ME: " & MaterialList(i).GetItemME

                            If RelicDecryptorText <> "" Then
                                If RelicDecryptorText.Contains("Intact") Or RelicDecryptorText.Contains("Malfunctioning") Or RelicDecryptorText.Contains("Wrecked") Then
                                    OutputString = OutputString & ", Relic: " & RelicDecryptorText
                                Else
                                    ' Decryptor
                                    OutputString = OutputString & ", Decryptor: " & RelicDecryptorText
                                End If
                            End If

                            OutputString = OutputString & ")"
                        End If
                    End If


                    If Not MatName.Contains("Data Interface") Then
                        OutputString = OutputString & " - " & FormatNumber(MaterialList(i).GetQuantity, 0) & vbCrLf
                    Else
                        OutputString = OutputString & vbCrLf
                    End If
                End If
SkipFormat:
            Next

            ' Add total volume and cost to end
            If Not IgnorePriceVolume Then

                OutputString = OutputString & DataInterfaces

                If ExportTextFormat = CSVDataExport Or ExportTextFormat = SSVDataExport Then
                    Separator = Trim(Separator) ' Remove space
                    OutputString = OutputString & vbCrLf & "Total Volume of Materials:" & Separator & CStr(TotalMaterialsVolume - DataInterfaceVolume) & Separator & "m3"
                    OutputString = OutputString & vbCrLf & "Total Cost of Materials:" & Separator & CStr(TotalMaterialsCost - DataInterfaceCost) & Separator & "ISK"
                Else
                    OutputString = OutputString & vbCrLf & "Total Volume of Materials: " & FormatNumber(TotalMaterialsVolume - DataInterfaceVolume, 2) & " m3"
                    OutputString = OutputString & vbCrLf & "Total Cost of Materials: " & FormatNumber(TotalMaterialsCost - DataInterfaceCost, 2) & " ISK"
                End If
            End If

            GetClipboardList = OutputString
        Else
            GetClipboardList = "No items in List" & vbCrLf
        End If

    End Function

    ' Returns the total cost of the material list
    Public Function GetTotalMaterialsCost() As Double
        Return TotalMaterialsCost
    End Function

    ' Returns the total volume of the matierals in the list
    Public Function GetTotalVolume() As Double
        Return TotalMaterialsVolume
    End Function

    ' Sorts the material list by quantity
    Private Sub SortListDesc(ByVal List() As Material, ByVal First As Integer, ByVal Last As Integer)
        Dim LowIndex As Integer
        Dim HighIndex As Integer
        Dim MidValue As Long

        ' Quicksort
        LowIndex = First
        HighIndex = Last
        MidValue = List((First + Last) \ 2).GetQuantity

        Do
            While List(LowIndex).GetQuantity > MidValue
                LowIndex = LowIndex + 1
            End While

            While List(HighIndex).GetQuantity < MidValue
                HighIndex = HighIndex - 1
            End While

            If LowIndex <= HighIndex Then
                Swap(LowIndex, HighIndex)
                LowIndex = LowIndex + 1
                HighIndex = HighIndex - 1
            End If
        Loop While LowIndex <= HighIndex

        If First < HighIndex Then
            SortListDesc(List, First, HighIndex)
        End If

        If LowIndex < Last Then
            SortListDesc(List, LowIndex, Last)
        End If

    End Sub

    ' This swaps the list values
    Private Sub Swap(ByRef IndexA As Integer, ByRef IndexB As Integer)
        Dim Temp As Material

        Temp = MaterialList(IndexA)
        MaterialList(IndexA) = MaterialList(IndexB)
        MaterialList(IndexB) = Temp

    End Sub

    ' Returns boolean comparison of two material lists
    Public Function MaterialListsEqual(ByVal List1 As Materials, ByVal List2 As Materials) As Boolean
        Dim i, j As Integer
        Dim Matlist1, Matlist2 As Material()
        Dim ItemFound As Boolean

        Matlist1 = List1.GetMaterialList
        Matlist2 = List2.GetMaterialList

        For i = 0 To Matlist1.Count - 1
            For j = 0 To Matlist2.Count - 1
                ' Looking for the item first, if not found then not equal
                ItemFound = False
                If Matlist1(i).GetMaterialName = Matlist2(j).GetMaterialName Then
                    ItemFound = True
                    If Matlist1(i).GetMaterialTypeID <> Matlist2(j).GetMaterialTypeID Then
                        Return False
                    End If
                    If Matlist1(i).GetMaterialGroup <> Matlist2(j).GetMaterialGroup Then
                        Return False
                    End If
                    If Matlist1(i).GetQuantity <> Matlist2(j).GetQuantity Then
                        Return False
                    End If
                    If Matlist1(i).GetVolume <> Matlist2(j).GetVolume Then
                        Return False
                    End If
                    If Matlist1(i).GetCostPerItem <> Matlist2(j).GetCostPerItem Then
                        Return False
                    End If
                    If Matlist1(i).GetItemME <> Matlist2(j).GetItemME Then
                        Return False
                    End If
                    If Matlist1(i).GetBuildItem <> Matlist2(j).GetBuildItem Then
                        Return False
                    End If
                    If Matlist1(i).GetTotalVolume <> Matlist2(j).GetTotalVolume Then
                        Return False
                    End If
                    If Matlist1(i).GetTotalCost <> Matlist2(j).GetTotalCost Then
                        Return False
                    End If
                    If Matlist1(i).GetItemType <> Matlist2(j).GetItemType Then
                        Return False
                    End If
                End If

                If ItemFound Then
                    ' Exit the loop if we found it
                    Exit For
                End If
            Next

            If Not ItemFound Then
                Return False
            End If

        Next

        Return True

    End Function

End Class
