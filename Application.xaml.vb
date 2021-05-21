Imports System.IO
Imports RequiredProductionClasses
Imports RequiredProductionClasses.Utilities

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

    Public Shared Event CurrentPOBaseChanged(sender As ProductOrder, e As EventArgs)
    Public Shared Event CurrentPOContentChanged(sender As Object, e As EventArgs)

    Private Sub Application_OnStartUp() Handles Me.Startup
        Product.GenerateProductList(ProductDirectory)

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

            If GetAllowedProducts.Contains(contentVar(1).Split(",")(4)) Then
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

End Class