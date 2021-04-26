Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Win32
Imports Product_Breakdown.Application
Imports Product_Breakdown.MainWindow
Imports Product_Breakdown.ProductOrder
Imports System.Windows.Xps

Class MainWindow

    Public Property ComponentBreakdownDocument As FixedDocumentSequence = Nothing

    Private Sub OnLoaded() Handles Me.Loaded
        'Add handler to needed controls:
        'AddHandler sldr_NumberOfDays.ValueChanged, AddressOf sldr_NumberOfDays_ValueChanged


        'Set control values to the setting values:
        cb_NumberOfDays.SelectedIndex = My.Settings.int_DaysToDisplay - 2

        'PO List Prep:
        'GeneratePOList()
        GeneratePODict()
        UpdatePOList()

        createWatcher_iSupplier()
    End Sub

    'Watches for file changes in the specified directory. Ensures that information regarding the directory in question is up-to-date.
    Public Sub createWatcher_iSupplier()
        Using watcher = New FileSystemWatcher(My.Settings.iSupplier_Default)
            watcher.NotifyFilter = NotifyFilters.LastWrite Or
                NotifyFilters.Size Or
                NotifyFilters.FileName

            AddHandler watcher.Changed, AddressOf watcher_ChangedFile
            AddHandler watcher.Created, AddressOf watcher_NewChangedFile
            AddHandler watcher.Deleted, AddressOf watcher_NewChangedFile
            AddHandler watcher.Renamed, AddressOf watcher_ChangedFile

            AddHandler Application.CurrentPOBaseChanged, AddressOf CurrentPO_BaseChanged
        End Using
    End Sub

    Public Function PrintComponentBreakDown(ProductFamily As String, ProductSubFamily As String, breakdownTotals As Dictionary(Of String, Integer)) As Boolean
        Try
            Dim outputDate As String = DateTime.Now.ToString("MM-dd-yyyy HH:mm")
            Dim outputName As String = String.Concat($"{My.Settings.BreakdownSaveDirectory}\{CurrentPO.PONumber}_{ProductFamily} {ProductSubFamily}", ".xps")

            Dim appXl As New Excel.Application
            Dim workbookXl As Excel.Workbook
            Dim worksheetXl As Excel.Worksheet

            workbookXl = appXl.Workbooks.Add($"{My.Settings.TemplateDirectory}{ProductFamily} {ProductSubFamily}.xlsx")
            worksheetXl = workbookXl.Sheets($"{ProductFamily} {ProductSubFamily}")

            With worksheetXl
                .Range("E1").Value = CurrentPO.PONumber
                .Range("E3").Value = CurrentPO.ScheduleNumber.Split("-")(2)

                For Each item In Schema
                    .Range(item.Value).Value = breakdownTotals(item.Key)
                Next
            End With

            Dim rootPath As String = $"{My.Settings.BreakdownSaveDirectory}\{CurrentPO.PONumber}\"
            Dim fileName As String = $"{CurrentPO.PONumber}_{ProductFamily} {ProductSubFamily}.xlsx"
            Dim tempPath As String = rootPath + fileName

            If Not My.Computer.FileSystem.DirectoryExists(rootPath) Then
                'Create the new directory if it doesn't already exist.
                My.Computer.FileSystem.CreateDirectory(rootPath)
            Else
                My.Computer.FileSystem.DeleteDirectory(rootPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
                My.Computer.FileSystem.CreateDirectory(rootPath)
            End If

            'Should save the file under a new name but as the same filetype.
            workbookXl.SaveAs(tempPath, FileFormat:=51, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

            'Exports the document as an XPS filetype.
            workbookXl.ExportAsFixedFormat(1, tempPath.Replace("xlsx", "xps"))

            Try
                worksheetXl.PrintOut()
            Catch ex As Exception
                MessageBox.Show("An error occurred during printing. Try again?", "Error!", MessageBoxButton.YesNo, MessageBoxImage.Error)
                If DialogResult = MessageBoxResult.Yes Then
                    Try
                        worksheetXl.PrintOut()
                    Catch ex1 As Exception

                    End Try
                End If
            End Try

            workbookXl.Close()
            appXl.Quit()

            ReleaseAll(appXl)
            ReleaseAll(workbookXl)

            Return True

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "",
        MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End Try
    End Function

    'Does not return the file just a boolean determining if the generation was succesful.
    Public Function GenerateComponentBreakDown(ProductFamily As String, ProductSubFamily As String, breakdownTotals As Dictionary(Of String, Integer), Optional PO As ProductOrder = Nothing) As String
        If PO Is Nothing Then
            PO = CurrentPO
        End If

        Try
            'Location Variables
            Dim rootPath As String = $"{My.Settings.BreakdownSaveDirectory}\{PO.PONumber}\"
            Dim fileName As String = $"{PO.PONumber}_{ProductFamily} {ProductSubFamily}.xlsx"
            Dim tempPath As String = rootPath + fileName

            'Check if the files already exist. Do not recreate them if it's the case.
            If Not My.Computer.FileSystem.DirectoryExists(rootPath) AndAlso Not My.Computer.FileSystem.FileExists(tempPath.Replace(".xlsx", ".xps")) Then

                'Delete any template copies beforehand.
                If File.Exists($"{My.Settings.TemplateDirectory}{ProductFamily} {ProductSubFamily}_{My.Computer.Name}.xlsx") Then
                    My.Computer.FileSystem.DeleteFile($"{My.Settings.TemplateDirectory}{ProductFamily} {ProductSubFamily}_{My.Computer.Name}.xlsx", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                'Copy the template file.
                My.Computer.FileSystem.CopyFile($"{My.Settings.TemplateDirectory}{ProductFamily} {ProductSubFamily}.xlsx", $"{My.Settings.TemplateDirectory}{ProductFamily} {ProductSubFamily}_{My.Computer.Name}.xlsx")

                'Excel specific declarations
                Dim appXl As New Excel.Application
                Dim workbookXl As Excel.Workbook
                Dim worksheetXl As Excel.Worksheet

                workbookXl = appXl.Workbooks.Add($"{My.Settings.TemplateDirectory}{ProductFamily} {ProductSubFamily}_{My.Computer.Name}.xlsx")
                worksheetXl = workbookXl.Sheets($"{ProductFamily} {ProductSubFamily}")

                With worksheetXl
                    .Range("E1").Value = PO.PONumber
                    .Range("E3").Value = PO.ScheduleNumber.Split("-")(2)

                    For Each item In Schema
                        .Range(item.Value).Value = breakdownTotals(item.Key)
                    Next
                End With

                If Not My.Computer.FileSystem.DirectoryExists(rootPath) Then
                    'Create the new directory if it doesn't already exist.
                    My.Computer.FileSystem.CreateDirectory(rootPath)
                Else
                    My.Computer.FileSystem.DeleteDirectory(rootPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    My.Computer.FileSystem.CreateDirectory(rootPath)
                End If

                'Should save the file under a new name but as the same filetype.
                workbookXl.SaveAs(tempPath, FileFormat:=51, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

                'Exports the document as an XPS filetype.
                workbookXl.ExportAsFixedFormat(1, tempPath.Replace("xlsx", "xps"))

                workbookXl.Close()
                appXl.Quit()

                ReleaseAll(appXl)
                ReleaseAll(workbookXl)
            End If

            Return tempPath.Replace(".xlsx", ".xps")

        Catch ex As Exception
            'Need to close the file.

            MessageBox.Show(ex.Message, "",
        MessageBoxButton.OK, MessageBoxImage.Information)
            Return Nothing
        End Try
    End Function

    Private Sub ReleaseAll(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub btn_FindPO_Click(sender As Object, e As RoutedEventArgs) Handles btn_FindPO.Click
        Dim openPicker As New OpenFileDialog With {.CheckFileExists = True, .InitialDirectory = My.Settings.iSupplier_Default, .Multiselect = False, .Title = "Point File Location", .Filter = "Text files (*.txt)|*.txt"}
        openPicker.ShowDialog()

        Dim PONumber As String = ""
        If openPicker.FileName <> "" AndAlso openPicker.FileName IsNot Nothing Then
            'Need to get the new location for the POs in order to copy the file.
            PONumber = openPicker.FileName.Split("\")(openPicker.FileName.Split("\").Count - 1).Replace(".txt", "")

            CurrentPO = New ProductOrder(PONumber)

            Dim breakdown As Dictionary(Of String, Integer) = CreateComponentBreakdown()

            PrintComponentBreakDown(GetFamilyInfo()(0), GetFamilyInfo()(1), breakdown)
        End If

        'If the a file is not chosen then the program doesn't continue executing.
    End Sub

    Private Function CreateComponentBreakdown(Optional PO As ProductOrder = Nothing) As Dictionary(Of String, Integer)
        Dim tempDict As New Dictionary(Of String, Integer)
        tempDict.Add("15H", 0)
        tempDict.Add("22.5H", 0)
        tempDict.Add("30H", 0)
        tempDict.Add("35H", 0)
        tempDict.Add("42.5H", 0)
        tempDict.Add("50H", 0)
        tempDict.Add("57.5H", 0)
        tempDict.Add("65H", 0)

        tempDict.Add("24W", 0)
        tempDict.Add("30W", 0)
        tempDict.Add("36W", 0)
        tempDict.Add("42W", 0)
        tempDict.Add("48W", 0)
        tempDict.Add("60W", 0)

        'If PO = nothing then we default to the current PO for the data.
        If PO Is Nothing Then
            PO = CurrentPO
        End If

        For Each product In PO.Products
            Try
                If ProductDirectory.Keys.Contains(product.Key) Then
                    Dim productInfo As String = ProductDirectory(product.Key).UI_ButtonDescription
                    Dim productClass As String = ProductDirectory(product.Key).Classification
                    Dim prodHeight As String = productInfo.Split("-")(0)
                    Dim prodWidth As String = productInfo.Split("-")(1).Split(" ")(0)

                    tempDict(prodHeight) += 2 * product.Value.QtyNeeded

                    If productClass = "Stacker" Then
                        tempDict(prodWidth) += CInt(product.Value.QtyNeeded)
                    ElseIf productClass = "Base Height" Then
                        If productInfo.Contains("35H") And ProductDirectory(product.Key).Family = "Terrace" Then
                            'Give 2 assuming Terrace.
                            tempDict(prodWidth) += product.Value.QtyNeeded * 2
                        Else
                            'Give 3.
                            tempDict(prodWidth) += product.Value.QtyNeeded * 3
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try
        Next

        Return tempDict
    End Function

    Private Function GetFamilyInfo(Optional Current As Boolean = True, Optional BucketFile As ProductOrder = Nothing) As List(Of String)
        Dim tempList As New List(Of String)

        If BucketFile Is Nothing And Current = True Then
            tempList.Add(ProductDirectory(CurrentPO.Products.Keys(0)).Family)
            tempList.Add(ProductDirectory(CurrentPO.Products.Keys(0)).SubFamily)
        ElseIf BucketFile IsNot Nothing And Current = False Then
            tempList.Add(ProductDirectory(BucketFile.Products.Keys(0)).Family)
            tempList.Add(ProductDirectory(BucketFile.Products.Keys(0)).SubFamily)
        End If

        Return tempList
    End Function

    Private Sub list_RecentPOs_SelectionChanged(sender As ListView, e As SelectionChangedEventArgs) Handles list_RecentPOs.SelectionChanged
        'Current PO changes based on the input received with the list view.
        'This will largely replace the button and open file dialog method and allow for the use of the UI.
        'The old method will still be used in case this method has an unknown error and doesn't allow for proper use or updating of the PO list.
        'This method also will only deal with a set number of POs.

        'Checks if the event was caused by an item being selected and not unselectd.
        If e.AddedItems IsNot Nothing AndAlso e.AddedItems.Count <> 0 Then
            Dim selectedItem As ListViewItem = e.AddedItems(0)
            'The current PO is only nothing until the fist list view item is selected.
            CurrentPO = New ProductOrder(selectedItem.Content)
        End If
    End Sub

    Public Sub watcher_NewChangedFile(sender As Object, e As FileSystemEventArgs)
        'Need re-add the list view contents.
        UpdatePOList()
    End Sub

    Public Sub watcher_ChangedFile(sender As Object, e As EventArgs)
        'Need to ensure that the current PO is up-to-date given the PO file changed is the same PO as the current PO.

    End Sub


    Private Structure Day_POObject
        Public ScheduleNumber As Date
        Public PONumber As String
    End Structure

    Private Sub UpdatePOList()
        'Clear list and then re-add all POs in need of re-adding.

        'Clear the list view items.
        'list_RecentPOs.Items.Clear()

        'You should only have the last 15 days in the list.

        Dim templist As New List(Of ListViewItem)

        For Each DayVar In DictOfPOs
            'Loop through each Day.
            If DictOfPOs.Keys.ToList.IndexOf(DayVar.Key) + 1 <= My.Settings.int_DaysToDisplay Then
                For Each POVar In DayVar.Value

                    'Add PO to the list view.

                    Dim tempString As String = POVar.Split("\")(POVar.Split("\").Count - 1).Replace(".txt", "")

                    'tempItem.PONumber = tempString
                    'tempItem.ScheduleNumber = CDate(DayVar.Key.Insert(4, "-").Insert(2, "-"))

                    Dim tempItem As New ListViewItem
                    tempItem.Content = tempString
                    tempItem.Tag = DayVar.Key.Insert(4, "-").Insert(2, "-")

                    Try
                        Dim tempPO As New ProductOrder(tempString)

                        If GetFamilyInfo(False, tempPO)(0) = "Terrace" Then
                            tempItem.Background = New SolidColorBrush(Color.FromRgb(255, 122, 231))
                        ElseIf GetFamilyInfo(False, tempPO)(0) = "Stride" Then
                            tempItem.Background = New SolidColorBrush(Color.FromRgb(51, 190, 255))
                        End If

                        templist.Add(tempItem)
                    Catch ex As Exception
                        'Does nothing and continues to the next product order.
                    End Try

                Next
            End If
        Next

        'Adding the content to the list view.
        list_RecentPOs.ItemsSource = templist

        Dim view As CollectionView = CollectionViewSource.GetDefaultView(list_RecentPOs.ItemsSource)
        Dim groupDescription As PropertyGroupDescription = New PropertyGroupDescription("Tag")
        view.GroupDescriptions.Add(groupDescription)
    End Sub

    Private Sub UpdatePOInformation(Optional isNothing As Boolean = False)
        If isNothing = False Then
            Me.lbl_DueDate.Content = $"Due Date: {CDate(CurrentPO.ScheduleNumber.Split("-")(2).Insert(4, "/").Insert(2, "/")).Subtract(New TimeSpan(5, 0, 0, 0)).ToShortDateString()}"
            Me.lbl_Family.Content = $"Family: {GetFamilyInfo()(0)}"
            Me.lbl_PONumber.Content = $"PO #: {CurrentPO.PONumber}"
            Me.lbl_ScheduleNumber.Content = $"Schedule #: {CurrentPO.ScheduleNumber}"

            If GetFamilyInfo()(0) = "Terrace" Then
                Me.lbl_Family.Background = New SolidColorBrush(Color.FromRgb(255, 122, 231))
            ElseIf GetFamilyInfo()(0) = "Stride" Then
                Me.lbl_Family.Background = New SolidColorBrush(Color.FromRgb(51, 190, 255))
            End If
        Else
            Me.lbl_DueDate.Content = "Due Date:"
            Me.lbl_Family.Content = "Family:"
            Me.lbl_Family.Background = New SolidColorBrush(Colors.Transparent)
            Me.lbl_PONumber.Content = "PO #: None Selected"
            Me.lbl_ScheduleNumber.Content = "Schedule #:"
        End If
    End Sub

    Private Sub CurrentPO_BaseChanged(sender As ProductOrder, e As EventArgs)
        'The current PO has at this point been changed to be a different PO.

        'Since there is new information, we need to update the information on the main UI window for the current PO.
        If sender.Products.Keys.Count = 0 Or sender.PONumber = Nothing Then
            UpdatePOInformation(True)
        Else
            UpdatePOInformation()
            FillBucketsTab(CurrentPO)
            ShowComponentBreakdown(CurrentPO)
            'PrintBuckets(CurrentPO)
        End If
    End Sub

    Private Sub CurrentPO_Changed()

    End Sub

    Private Sub btn_PrintBreakdown_Click(sender As Object, e As EventArgs) Handles btn_PrintBreakdown.Click
        Dim breakDown As New Dictionary(Of String, Integer)
        breakDown = CreateComponentBreakdown()

        PrintComponentBreakDown(GetFamilyInfo()(0), GetFamilyInfo()(1), breakDown)
    End Sub

    Private Sub ShowComponentBreakdown(PO As ProductOrder)
        'Generate the Component Breakdown.

        Dim breakdown As Dictionary(Of String, Integer) = CreateComponentBreakdown(PO)
        Dim familyInfo As List(Of String) = GetFamilyInfo(False, PO)

        Dim breakdownSaveLocation As String = GenerateComponentBreakDown(familyInfo(0), familyInfo(1), breakdown)

        'Set the XPS document as the document viewer's document.

        If File.Exists(breakdownSaveLocation) Then
            Try
                Dim xpsDocument As New Xps.Packaging.XpsDocument(breakdownSaveLocation, FileAccess.Read)
                Dim fixedDocSeq As FixedDocumentSequence = xpsDocument.GetFixedDocumentSequence()

                doc_ComponentBreakdown.Document = fixedDocSeq
            Catch ex As UnauthorizedAccessException
                MessageBox.Show($"HResult: {ex.HResult}{vbNewLine}Message: {ex.Message}{vbNewLine}", "Unable To Access File.", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        Else
            doc_ComponentBreakdown.Document = Nothing
        End If

    End Sub

    Private Sub FillBucketsTab(PO As ProductOrder)
        listView_Buckets.ItemsSource = PO.BucketizeProducts

        Dim view As CollectionView = CollectionViewSource.GetDefaultView(listView_Buckets.ItemsSource)
        Dim groupDescription As PropertyGroupDescription = New PropertyGroupDescription("BucketNumber")
        view.GroupDescriptions.Add(groupDescription)

        'Counts:
        Dim RogueCount As Integer = PO.RogueBuckets.Count
        Dim BulkCount As Integer = PO.BulkPacks.Count

        'Fill header information.
        lbl_BucketsHeader.Content = $"Bulk Packs : {PO.BulkPacks.Count} | Singles : {PO.RogueBuckets.Count} | % Bulk Packs : {Math.Round((PO.BulkPacks.Count / PO.Buckets.Count) * 100, 0)}%"
    End Sub

    Private Sub PrintRogueBuckets(PO As ProductOrder)

        Try
            Dim outputList As New Dictionary(Of String, Integer)
            Dim outputList_Width As New Dictionary(Of String, Dictionary(Of String, Integer))

            For Each BucketNum In PO.RogueBuckets
                Dim tempBucket As Bucket = PO.Buckets(BucketNum)

                For Each productVar In tempBucket.Products
                    'Distinguish width.
                    Dim widthVar As String = ProductDirectory(productVar.Key).UI_ButtonDescription.Split(" ")(0).Split("-")(1)

                    If outputList_Width.ContainsKey(widthVar) Then
                        'Already exists.
                        If outputList_Width(widthVar).ContainsKey(productVar.Key) Then

                            outputList_Width(widthVar)(productVar.Key) += productVar.Value.QtyNeeded

                        Else

                            outputList_Width(widthVar).Add(productVar.Key, productVar.Value.QtyNeeded)

                        End If

                    Else
                        'Does not exist.
                        outputList_Width.Add(widthVar, New Dictionary(Of String, Integer))
                        outputList_Width(widthVar).Add(productVar.Key, productVar.Value.QtyNeeded)
                    End If
                Next
            Next

            'Sort by width and add to the workbook.
            For Each widthVar In outputList_Width
                Dim appXl As New Excel.Application
                Dim workbookXl As Excel.Workbook
                Dim worksheetXl As Excel.Worksheet

                workbookXl = appXl.Workbooks.Add($"{My.Settings.TemplateDirectory}Single Pack Template.xlsx")
                worksheetXl = workbookXl.Sheets("Sheet1")

                With worksheetXl
                    'Add the header.
                    .Range($"B1").Value = PO.ScheduleNumber.Remove(0, PO.ScheduleNumber.Split("-")(0).Count + 1)

                    'Add each collection item to the worksheet.
                    Dim index As Integer = 2
                    For Each item In widthVar.Value
                        'Set up the cells:
                        .Range($"B{index}").BorderAround2(Weight:=Excel.XlBorderWeight.xlThin)
                        .Range($"B{index}").Value = PO.Products(item.Key).Description
                        .Range($"C{index}").BorderAround2(Weight:=Excel.XlBorderWeight.xlThin)
                        .Range($"C{index}").Value = item.Key    'Part #
                        .Range($"D{index}").BorderAround2(Weight:=Excel.XlBorderWeight.xlThin)
                        .Range($"D{index}").Value = item.Value  'Total Count

                        index += 1
                    Next

                    'Add the footer.
                    .Range($"B{index}").Value = "Grand Total"
                    .Range($"B{index}").BorderAround2(Weight:=Excel.XlBorderWeight.xlMedium)
                    .Range($"C{index}").BorderAround2(Weight:=Excel.XlBorderWeight.xlMedium)
                    .Range($"D{index}").Value = widthVar.Value.Values.Sum()
                    .Range($"D{index}").BorderAround2(Weight:=Excel.XlBorderWeight.xlMedium)
                End With

                Try
                    worksheetXl.PrintOut()
                Catch ex As Exception
                    MessageBox.Show("An error occurred during printing. Try again?", "Error!", MessageBoxButton.YesNo, MessageBoxImage.Error)
                    If DialogResult = MessageBoxResult.Yes Then
                        Try
                            worksheetXl.PrintOut()
                        Catch ex1 As Exception

                        End Try
                    End If
                End Try

                workbookXl.Close(SaveChanges:=False)
                appXl.Quit()

                ReleaseAll(appXl)
                ReleaseAll(workbookXl)
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "",
        MessageBoxButton.OK, MessageBoxImage.Information)
        End Try
    End Sub

    Private Sub tabC_MainTabControl_SelectionChanged(sender As TabControl, e As SelectionChangedEventArgs) Handles tabC_MainTabControl.SelectionChanged
        If tabC_MainTabControl.SelectedItem Is tab_ComponentBreakdown Then
            Me.Width = Me.MinWidth + 250
            Me.Height = Me.MinHeight + 400
        ElseIf tabC_MainTabControl.SelectedItem IsNot tab_ComponentBreakdown Then
            Me.Width = Me.MinWidth
            Me.Height = Me.MinHeight
        End If
    End Sub

    Private Sub btn_PrintSingles_Click(sender As Object, e As RoutedEventArgs) Handles btn_PrintSingles.Click
        If CurrentPO IsNot Nothing Then
            PrintRogueBuckets(CurrentPO)
        End If
    End Sub

    Private Sub cb_NumberOfDays_SelectionChanged(sender As ComboBox, e As SelectionChangedEventArgs) Handles cb_NumberOfDays.SelectionChanged
        'Combobox does not distinguish what POs are brought in, only how many general POs are brought in.

        'Set the count in settings.
        My.Settings.int_DaysToDisplay = CInt(sender.SelectedItem.Content)

        My.Settings.Save()

        'Update the list of POs that are currently displayed.
        UpdatePOList()
    End Sub
End Class
