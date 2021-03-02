Imports Microsoft.Office.Interop
Imports Microsoft.Win32
Imports Product_Breakdown.Application
Imports Product_Breakdown.BucketPickFile
Imports Product_Breakdown.MainWindow

Class MainWindow

    Public Function PrintBreakDown(ProductFamily As String, ProductSubFamily As String, breakdownTotals As Dictionary(Of String, Integer)) As Boolean
        Try
            Dim outputDate As String = DateTime.Now.ToString("MM-dd-yyyy HH:mm")
            Dim outputName As String = String.Concat($"{My.Settings.BreakdownSaveDirectory}\{CurrentPO.PONumber}_{ProductFamily} {ProductSubFamily}", ".xps")

            Dim appXl As New Excel.Application
            Dim workbookXl As Excel.Workbook
            Dim worksheetXl As Excel.Worksheet

            workbookXl = appXl.Workbooks.Add(My.Settings.TemplateDirectory)
            worksheetXl = workbookXl.Sheets($"{ProductFamily} {ProductSubFamily}")

            With worksheetXl
                .Range("E1").Value = CurrentPO.PONumber
                .Range("E3").Value = CurrentPO.ScheduleNumber.Split("-")(2)

                For Each item In Schema
                    .Range(item.Value).Value = breakdownTotals(item.Key)
                Next
            End With


            My.Computer.FileSystem.CreateDirectory($"{My.Settings.BreakdownSaveDirectory}\{CurrentPO}_{ProductFamily} {ProductSubFamily}")
            'Should save the file under a new name but as the same filetype.
            workbookXl.SaveAs(String.Concat($"{My.Settings.BreakdownSaveDirectory}\{CurrentPO.PONumber}_{ProductFamily} {ProductSubFamily}"), FileFormat:=51)

            'Exports the document as an XPS filetype.
            workbookXl.ExportAsFixedFormat(1, outputName)

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

        Dim fileLocation As String = ""
        If openPicker.FileName <> "" And openPicker.FileName IsNot Nothing Then
            'Need to get the new location for the POs in order to copy the file.
            Dim newLocation As String = $"{My.Settings.iSupplier_CopiedLocation}{openPicker.FileName.Split("\")(openPicker.FileName.Split("\").Count - 1)}"
            My.Computer.FileSystem.CopyFile(openPicker.FileName, newLocation, True)
            fileLocation = newLocation
        End If

        CurrentPO = New BucketPickFile(fileLocation)

        Dim breakdown As Dictionary(Of String, Integer) = CreateBreakdown()

        PrintBreakDown(GetFamilyInfo()(0), GetFamilyInfo()(1), breakdown)
    End Sub

    Private Function CreateBreakdown() As Dictionary(Of String, Integer)
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

        For Each product In CurrentPO.ProductionNumbers
            Try
                If ProductDirectory.Keys.Contains(product.Key) Then
                    Dim productInfo As String = ProductDirectory(product.Key).UI_ButtonDescription
                    Dim productClass As String = ProductDirectory(product.Key).Classification
                    Dim prodHeight As String = productInfo.Split("-")(0)
                    Dim prodWidth As String = productInfo.Split("-")(1).Split(" ")(0)

                    tempDict(prodHeight) += 2 * product.Value(0)

                    If productClass = "Stacker" Then
                        tempDict(prodWidth) += CInt(product.Value(0))
                    ElseIf productClass = "Base Height" Then
                        If productInfo.Contains("35H") And ProductDirectory(product.Key).Family = "Terrace" Then
                            'Give 2 assuming Terrace.
                            tempDict(prodWidth) += product.Value(0) * 2
                        Else
                            'Give 3.
                            tempDict(prodWidth) += product.Value(0) * 3
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try
        Next

        Return tempDict
    End Function

    Private Function GetFamilyInfo() As List(Of String)
        Dim tempList As New List(Of String)
        tempList.Add(ProductDirectory(CurrentPO.ProductionNumbers.Keys(0)).Family)
        tempList.Add(ProductDirectory(CurrentPO.ProductionNumbers.Keys(0)).SubFamily)

        Return tempList
    End Function
End Class
