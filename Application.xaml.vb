Imports System.IO

Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

    Private Shared _CurrentPO As ProductOrder
    Public Shared Property CurrentPO() As ProductOrder
        Get
            Return _CurrentPO
        End Get
        Set(ByVal value As ProductOrder)
            _CurrentPO = value
            RaiseEvent CurrentPOBaseChanged(_CurrentPO, Nothing)
        End Set
    End Property

    Public Shared ProductDirectory As New Dictionary(Of String, Product)

    Public Shared Schema As New Dictionary(Of String, String)

    Public Shared ProductFamily As String
    Public Shared ProductSubFamily As String

    Public Shared ListOfPOs As New List(Of String)
    Public Shared DictOfPOs As New Dictionary(Of String, List(Of String))

    'Holds all part numbers for the products that are allowed for use in this program.
    Public Shared AllowedProductRegistry As New List(Of String)

    Public Shared Event CurrentPOBaseChanged(sender As ProductOrder, e As EventArgs)
    Public Shared Event CurrentPOContentChanged(sender As Object, e As EventArgs)

    Private Sub Application_OnStartUp() Handles Me.Startup
        Product.GenerateProductList()

        Schema.Add("15H", "C9")
        Schema.Add("22.5H", "C11")
        Schema.Add("30H", "C13")

        Schema.Add("35H", "C17")
        Schema.Add("42.5H", "C19")
        Schema.Add("50H", "C21")
        Schema.Add("57.5H", "C23")
        Schema.Add("65H", "C25")

        Schema.Add("24W", "F7")
        Schema.Add("30W", "F9")
        Schema.Add("36W", "F11")
        Schema.Add("42W", "F13")
        Schema.Add("48W", "F15")
        Schema.Add("60W", "F17")
    End Sub

    Public Shared Sub GeneratePOList()
        'Compile the list of POs from the iSupplier directory.
        For Each fileVar In My.Computer.FileSystem.GetFiles(My.Settings.iSupplier_Default, FileIO.SearchOption.SearchTopLevelOnly)

            If DateTime.Now.Subtract(FileDateTime(fileVar)).TotalDays < My.Settings.int_DaysToDisplay AndAlso isOnlyNumeric(fileVar.Split("\").Reverse()(0).Replace(".txt", "").Replace("-", "")) = True Then


                If Not ListOfPOs.Contains(fileVar) AndAlso isQualified(fileVar.Split("\")(fileVar.Split("\").Count - 1).Replace(".txt", "")) Then
                    ListOfPOs.Add(fileVar)
                End If

            End If

        Next
    End Sub

    Public Shared Sub GeneratePODict()
        'Compile top-level list.
        Dim ListOfPOFiles As List(Of String) = My.Computer.FileSystem.GetFiles(My.Settings.iSupplier_Default, FileIO.SearchOption.SearchTopLevelOnly).ToList()

        'Reverse the list so the newest POs are on the top.
        ListOfPOFiles.Reverse()



        'ListOfPOFiles = ListOfPOs.OrderBy(Function(x) x.Split("-")(1)).ToList()

        Dim index As Integer = 0
        While DictOfPOs.Keys.Count < 10 AndAlso index < ListOfPOFiles.Count

            'Read the lines of the file.
            Dim contentVar As List(Of String) ' = File.ReadAllText(ListOfPOFiles(index)).Split(vbNewLine).ToList()
            Try

                contentVar = File.ReadAllText(ListOfPOFiles(index)).Split(vbNewLine).ToList()

            Catch ex As IOException

                'File is open in another process.
                'Copy the file into the secondary directory.
                Dim newloc As String = $"{My.Settings.iSupplier_CopiedLocation}{ListOfPOFiles(index).Split("\").Last()}"

                If Not File.Exists(newloc) Then
                    My.Computer.FileSystem.CopyFile(ListOfPOFiles(index), newloc)
                End If

                'Read lines in from new location.
                contentVar = File.ReadAllText($"{My.Settings.iSupplier_CopiedLocation}{ListOfPOFiles(index).Split("\").Last()}").Split(vbNewLine).ToList()
            End Try

            If My.Settings.AllowedProductRegistry.Contains(contentVar(1).Split(",")(4)) Then
                'Get the Schedule Number.
                Dim scheduleNumber As String = contentVar(1).Split(",")(1).Split("-")(2)

                'Add the schedule number and PO to the dictionary.
                If DictOfPOs.Keys.Contains(scheduleNumber) Then
                    'Already present.

                    DictOfPOs(scheduleNumber).Add(ListOfPOFiles(index))
                Else
                    'Not already present.

                    DictOfPOs.Add(scheduleNumber, New List(Of String))
                    DictOfPOs(scheduleNumber).Add(ListOfPOFiles(index))
                End If
            End If

            index += 1
        End While
    End Sub

    Public Shared Function isQualified(PONumber As String) As Boolean
        'Verify the PO is for qualified products.
        'Qualified Products are products that are programed in.

        Dim tempPO As New ProductOrder(PONumber)

        'Verify part numbers are apart of the allowed registry.
        If My.Settings.AllowedProductRegistry.Contains(tempPO.Products.Keys(0)) Then
            tempPO = Nothing
            Return True
        End If

        Return False
    End Function

    'Delete after use.
    Private Sub updateProducts()
        Dim tempDict_T As New Dictionary(Of String, String)
        tempDict_T.Add("15H", "3790569000")
        tempDict_T.Add("22.5H", "3790569100")
        tempDict_T.Add("30H", "3790569200")
        tempDict_T.Add("35H", "3790569300")
        tempDict_T.Add("42.5H", "3790569400")
        tempDict_T.Add("50H", "3790569500")
        tempDict_T.Add("57.5H", "3790569600")
        tempDict_T.Add("65H", "3790569700")

        tempDict_T.Add("24W", "3760983200")
        tempDict_T.Add("30W", "3760983300")
        tempDict_T.Add("36W", "3760983400")
        tempDict_T.Add("42W", "3760983500")
        tempDict_T.Add("48W", "3760983600")
        tempDict_T.Add("60W", "3760983700")

        Dim tempDict_S As New Dictionary(Of String, String)
        tempDict_S.Add("15H", "3760539100")
        tempDict_S.Add("22.5H", "3760540000")
        tempDict_S.Add("30H", "3760539200")
        tempDict_S.Add("35H", "3760539300")
        tempDict_S.Add("42.5H", "3760539400")
        tempDict_S.Add("50H", "3760539500")
        tempDict_S.Add("57.5H", "3760539600")
        tempDict_S.Add("65H", "3760539700")

        tempDict_S.Add("24W", "3760983200")
        tempDict_S.Add("30W", "3760983300")
        tempDict_S.Add("36W", "3760983400")
        tempDict_S.Add("42W", "3760983500")
        tempDict_S.Add("48W", "3760983600")
        tempDict_S.Add("60W", "3760983700")


        For Each productVar As Product In ProductDirectory.Values
            If productVar.SubFamily = "Frame" Then
                Dim tempTupleList As New List(Of String)
                Dim tempTupleList1 As New List(Of String)

                If My.Settings.AllowedProductRegistry.Contains(productVar.PartNumber) Then
                    Dim productInfo As String = productVar.UI_ButtonDescription
                    Dim productClass As String = productVar.Classification
                    Dim prodHeight As String = productInfo.Split("-")(0)
                    Dim prodWidth As String = productInfo.Split("-")(1).Split(" ")(0)

                    If productVar.Family = "Terrace" Then
                        tempTupleList.Add(tempDict_T(prodHeight))
                        tempTupleList.Add("2")

                        If productClass = "Stacker" Then
                            tempTupleList1.Add(tempDict_T(prodWidth))
                            tempTupleList1.Add("1")
                        ElseIf productClass = "Base Height" Then
                            If productInfo.Contains("35H") Then
                                'Give 2 assuming Terrace.
                                tempTupleList1.Add(tempDict_T(prodWidth))
                                tempTupleList1.Add("2")
                            Else
                                'Give 3.
                                tempTupleList1.Add(tempDict_T(prodWidth))
                                tempTupleList1.Add("3")
                            End If
                        End If

                        productVar.ProductAssemblyInformation.ChildProduct1 = New Tuple(Of String, Integer)(tempTupleList(0), CInt(tempTupleList(1)))

                        productVar.ProductAssemblyInformation.ChildProduct2 = New Tuple(Of String, Integer)(tempTupleList1(0), CInt(tempTupleList1(1)))

                        productVar.ProductAssemblyInformation.ChildProductsCount = 2
                    ElseIf productVar.Family = "Stride" Then
                        tempTupleList.Add(tempDict_S(prodHeight))
                        tempTupleList.Add("2")

                        If productClass = "Stacker" Then
                            tempTupleList1.Add(tempDict_S(prodWidth))
                            tempTupleList1.Add("1")
                        ElseIf productClass = "Base Height" Then
                            If productInfo.Contains("35H") Then
                                'Give 2 assuming Terrace.
                                tempTupleList1.Add(tempDict_S(prodWidth))
                                tempTupleList1.Add("2")
                            Else
                                'Give 3.
                                tempTupleList1.Add(tempDict_S(prodWidth))
                                tempTupleList1.Add("3")
                            End If
                        End If

                        productVar.ProductAssemblyInformation.ChildProduct1 = New Tuple(Of String, Integer)(tempTupleList(0), CInt(tempTupleList(1)))

                        productVar.ProductAssemblyInformation.ChildProduct2 = New Tuple(Of String, Integer)(tempTupleList1(0), CInt(tempTupleList1(1)))

                        productVar.ProductAssemblyInformation.ChildProductsCount = 2
                    End If
                End If
            ElseIf productVar.SubFamily <> "Frame" Then
                productVar.ProductAssemblyInformation.ChildProduct1 = New Tuple(Of String, Integer)(productVar.PartNumber, 1)
                productVar.ProductAssemblyInformation.ChildProductsCount = 1
            End If

            productVar.SaveProduct()
        Next
    End Sub

    'Utilities
    Public Shared Function isNumeric(inputString As String) As Boolean
        Dim characterDictionary As New List(Of String) From {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
        For Each character In inputString
            If characterDictionary.Contains(character) Then
                'If the string contains even a single number we consider it numeric.
                Return True
            End If
        Next

        'Only gets here if the string contains no numbers.
        Return False
    End Function

    Public Shared Function isOnlyNumeric(inputString As String) As Boolean
        Try
            Dim testVar As Double = CDbl(inputString)

            'inputString is convertable to a double therefore it is only numeric.
            Return True
        Catch ex As Exception
            'inputString failed the converstion and is therefore not just numeric.
            Return False
        End Try
    End Function
End Class