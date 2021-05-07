Imports System.IO
Imports Product_Breakdown.Application
Imports Newtonsoft.Json

Public Class ProductOrder

    'For any PO that contains all unregistered POs, their buckets will be shown as empty.

    'Variable/Property Declarations:
    Public Property ScheduleNumber As String
    Public Property PONumber As String
    Public Property OriginFile As String
    Public Property ScheduleNumber_Date As Date
    Public Property DueDate As Date

    'Updated everytime a new line is read and the current product is registered.
    Public Property TotalNeeded As Integer
    Public Property TotalProduced As Integer


    'Collections:
    Public Buckets As Dictionary(Of String, Bucket)
    Public LineItems As Dictionary(Of String, LineItem)
    Public Products As Dictionary(Of String, ProductEntry)
    Public Inputs As Dictionary(Of String, UserInput)

    'Holds the bucket number for all bulkpacks
    Public BulkPacks As List(Of String)
    'Holds the bucket number for all buckets that are not bulkpacks.
    Public RogueBuckets As List(Of String)

    'Rogue Buckets are any buckets that:
    '   - Have less than 14 total parts in it.
    '   - Have more than 1 part number attached.


    'Structure Definition for a line item:
    'Does not get updated when products are produced.
    Public Structure LineItem

        'Properties that make up a Line Item
        Public BucketNumber As String
        Public ScheduleNumber As String
        Public Aliass As String   'Named with an extra "s" to ensure as close to the original name as possible.
        Public BaseModel As String
        Public PartNumber As String
        Public PartDescription As String
        Public PartQty As Integer
        Public Seq As Integer  'Identifier
        Public PONumber As String
        Public ShipTo As String
        Public POLineNumber As Integer
        Public ShipmentNumber As Integer
        Public NeedByDate As String
        Public BucketizedFlag As Boolean
    End Structure


    'Structure Definition for a Bucket:
    'Needs to be updated as products are produced.
    Public Class Bucket

        Public BucketNumber As String
        Public Products As Dictionary(Of String, ProductEntry) 'Stores counts for the products within the bucket.
        Public LineItems As List(Of String) 'All lineitems that make up the bucket.
        Public BucketNeeded As Integer
        Public BucketProduced As Integer

        'Identifies the bucket as a bulkpack or rogue bucket.
        Public Property isBulkPack As Boolean

        'BucketA should always be the bigger of the two buckets.
        Public Overloads Shared Operator +(BucketA As Bucket, BucketB As Bucket)
            'BucketB will never have more than one item in it.
            '   -> Only made up of one line.

            'Return variable declaration:
            Dim returnBucket As Bucket
            returnBucket.BucketNumber = BucketA.BucketNumber
            returnBucket.Products = BucketA.Products
            returnBucket.LineItems = BucketA.LineItems

            'Verify bucket numbers are the same.
            If BucketB.BucketNumber <> BucketA.BucketNumber Or BucketA.LineItems.Contains(BucketB.LineItems.ElementAt(0)) Then
                Throw New ArgumentException("Buckets do not have the same bucket number or is a duplicate bucket.")
                'Will not continue on if they are different buckets.
            End If

            'Verify the incoming bucket does not have more than one product present.
            If BucketB.Products.Count > 1 Then
                Throw New ArgumentException("Bucket B does has too many line items.")

                'To-Do
            End If

            'Executes only if the product is registered.
            'Determines if the product is a registered product.
            If My.Settings.AllowedProductRegistry.Contains(BucketB.Products.ElementAt(0).Key) Then

                'Check if the product is already in the product dictionary.
                If BucketA.Products.ContainsKey(BucketB.Products.ElementAt(0).Key) Then
                    Dim PartNum As String = BucketB.Products.ElementAt(0).Key

                    'Combine the two product variables.
                    returnBucket.Products(PartNum) = BucketA.Products(PartNum) + BucketB.Products.ElementAt(0).Value

                Else
                    'Add the product to the dictionary.
                    returnBucket.Products.Add(BucketB.Products.ElementAt(0).Key, BucketB.Products.ElementAt(0).Value)
                End If

            End If

            'Add the line item to the bucket.
            returnBucket.LineItems.Add(BucketB.LineItems.ElementAt(0))

            returnBucket.BucketNeeded = BucketA.BucketNeeded + BucketB.BucketNeeded
            returnBucket.BucketProduced = BucketA.BucketProduced + BucketB.BucketProduced

            Return returnBucket
        End Operator

        Public Function Reset() As Integer
            'Temporary integer to store the count produced before reset.
            Dim tempInt As Integer = BucketProduced

            'Loop through the products and reset the produced counts.
            For Each productVar As ProductEntry In Products.Values
                'Reset produced count.
                productVar.QtyProduced = 0
            Next

            'Return the number of produced parts before the reset occured.
            Return tempInt
        End Function

        Public Sub New()
            Me.BucketNeeded = 0
            Me.BucketProduced = 0
            Me.BucketNumber = Nothing
            Me.LineItems = New List(Of String)
            Me.Products = New Dictionary(Of String, ProductEntry)
            Me.isBulkPack = False
        End Sub

        Public Sub New(BucketNumber As String)
            Me.BucketNeeded = 0
            Me.BucketProduced = 0
            Me.BucketNumber = BucketNumber
            Me.LineItems = New List(Of String)
            Me.Products = New Dictionary(Of String, ProductEntry)
            Me.isBulkPack = False
        End Sub
    End Class


    'Structure Definition for a UserInput:
    Public Structure UserInput

        'Deals with any input made by the user 
    End Structure

    'Structure definition for a product entry in the PO:
    'Needs to be updated when a product is produced.
    Public Class ProductEntry
        Public PartNumber As String
        Public Description As String
        Public QtyNeeded As Integer
        Public QtyProduced As Integer

        Public Overloads Shared Operator +(ProdA As ProductEntry, ProdB As ProductEntry) As ProductEntry
            'ProdA is the base product.

            'Check if the part numbers are the same.
            If ProdA.PartNumber <> ProdB.PartNumber Then
                'Not the same products. Need to throw an error.
                Throw New ArgumentException("Passed products are not the same product. Cannot add two different products.")
            End If

            'Setup return variable.
            Dim returnProduct As ProductEntry
            returnProduct.PartNumber = ProdA.PartNumber
            returnProduct.Description = ProdA.Description

            'Calculations:
            returnProduct.QtyNeeded = ProdA.QtyNeeded + ProdB.QtyNeeded
            returnProduct.QtyProduced = ProdA.QtyProduced + ProdB.QtyProduced

            Return returnProduct
        End Operator

        'Adds the identifier for the bucket it belongs to and creates the new type.
        Public Function Bucketize(bucketNumber As String) As BucketizedProductEntry
            Dim returnVar As BucketizedProductEntry

            returnVar.BucketNumber = bucketNumber
            returnVar.PartNumber = PartNumber
            returnVar.Description = Description
            returnVar.QtyNeeded = QtyNeeded
            returnVar.QtyProduced = QtyProduced

            Return returnVar
        End Function

        Public Sub New()
            Me.Description = Nothing
            Me.PartNumber = Nothing
            Me.QtyNeeded = 0
            Me.QtyProduced = 0
        End Sub

        Public Sub New(PartNumber As String)
            Me.Description = Nothing
            Me.PartNumber = PartNumber
            Me.QtyNeeded = 0
            Me.QtyProduced = 0
        End Sub
    End Class

    'Structure definition for a bucketized product entry in the PO:
    Public Structure BucketizedProductEntry
        Public Property PartNumber As String
        Public Property Description As String
        Public Property QtyNeeded As Integer
        Public Property QtyProduced As Integer
        Public Property BucketNumber As String
    End Structure

    Public Sub New(PONumber As String)
        'Initialize collections:
        Me.Buckets = New Dictionary(Of String, Bucket)
        Me.Products = New Dictionary(Of String, ProductEntry)
        Me.LineItems = New Dictionary(Of String, LineItem)
        Me.Inputs = New Dictionary(Of String, UserInput)
        Me.BulkPacks = New List(Of String)
        Me.RogueBuckets = New List(Of String)

        'Made to be ambiguous.
        'LoadFile($"{My.Settings.iSupplier_Default}{PONumber}.txt")
        'Me.OriginFile = $"{My.Settings.iSupplier_Default}{PONumber}"

        If File.Exists($"{My.Settings.POSaveDirectory}{PONumber}.json") Then
            Load(PONumber)

            If Me.DueDate = Nothing Then
                'Set DueDate:
                Me.DueDate = CDate(Me.LineItems(1).NeedByDate)
            End If
        Else
            LoadFile($"{My.Settings.iSupplier_Default}{PONumber}.txt")
            Me.OriginFile = $"{My.Settings.iSupplier_Default}{PONumber}"

            'Set DueDate:
            Me.DueDate = CDate(Me.LineItems(1).NeedByDate)

            Me.Save()
        End If

        'Determines if the buckets are bulkpacks or not.
        For Each bucketVar As Bucket In Buckets.Values

            'Checks if bucket has more than 1 part number and has 14 parts included in the bucket.
            If bucketVar.Products.Count = 1 Then
                'AndAlso ((bucketVar.BucketNeeded = 14 And Products.ElementAt(0).Value.Description.Contains("STRIDE WELDED FRAME")) Or (bucketVar.BucketNeeded = 16 And Products.ElementAt(0).Value.Description.Contains("TERR WELDED FRAME"))) Then
                bucketVar.isBulkPack = True

                'Add to the collection.
                If Not BulkPacks.Contains(bucketVar.BucketNumber) Then
                    BulkPacks.Add(bucketVar.BucketNumber)
                End If
            Else
                bucketVar.isBulkPack = False

                'Add to the collection.
                If Not RogueBuckets.Contains(bucketVar.BucketNumber) Then
                    RogueBuckets.Add(bucketVar.BucketNumber)
                End If
            End If
        Next

        'Setting the schedule number date property.
        Me.ScheduleNumber_Date = CDate(Me.ScheduleNumber.Split("-").Last.Insert(4, "-").Insert(2, "-"))

    End Sub



    <JsonConstructor>
    Public Sub New(PONumber As String, Buckets As Dictionary(Of String, Bucket), Products As Dictionary(Of String, ProductEntry), LineItems As Dictionary(Of String, LineItem), Inputs As Dictionary(Of String, UserInput), File As String, TotalNeeded As Integer, TotalProduced As Integer, ScheduleNumber As String, DueDate As Date)

        Me.PONumber = PONumber
        Me.ScheduleNumber = ScheduleNumber
        Me.OriginFile = File
        Me.TotalNeeded = TotalNeeded
        Me.TotalProduced = TotalProduced
        Me.Buckets = Buckets
        Me.Products = Products
        Me.LineItems = LineItems
        Me.Inputs = Inputs
        Me.DueDate = DueDate

    End Sub

    Public Function LoadFile(Path As String) As Boolean
        'Create new string for the path of the PO from the saved general path in app.settings and the PONumber variable.
        Try
            If File.Exists(Path) Then

                'Determines what order of columns the file has bases on the first row.
                Dim FileSchema As New Dictionary(Of String, Integer)

                'Start parsing the file.
                Using inFile As StreamReader = File.OpenText(Path)
                    Dim tempLineVar As List(Of String) = inFile.ReadLine.Split(",").ToList()
                    Dim headerVar = 0

                    FileSchema = IdentifyFileSchema(tempLineVar)

                    'Repeats until the end of the file.
                    Do While Not inFile.EndOfStream AndAlso FileSchema IsNot Nothing
                        Dim curLine As List(Of String) = Nothing
                        If tempLineVar.Count > 1 Then
                            curLine = inFile.ReadLine().Split(",").ToList()
                        Else
                            curLine = inFile.ReadLine().Split(vbTab).ToList()
                        End If
                        'Add statements to check against other forms of deleminators.

                        If headerVar = 0 Then
                            Me.ScheduleNumber = curLine(1).Replace($"-{curLine(FileSchema("SCHEDULE NUMBER")).Split("-")(3)}", "")
                            Me.PONumber = curLine(FileSchema("PO NUMBER"))
                            'Add 1 to headerVar as to not repeat the process a second time.
                            headerVar = 1
                        End If


                        'Me.LineItems
                        'Get the line item and add it to the dictionary.
                        Dim curLineItem As LineItem
                        Try
                            curLineItem.Aliass = curLine(FileSchema("ALIAS"))
                            curLineItem.BaseModel = curLine(FileSchema("BASE MODEL"))
                            curLineItem.BucketNumber = curLine(FileSchema("BUCKET NUMBER"))
                            curLineItem.NeedByDate = curLine(FileSchema("NEED BY DATE"))
                            curLineItem.PartDescription = curLine(FileSchema("PART DESCRIPTION"))
                            curLineItem.PartNumber = curLine(FileSchema("PART NUMBER"))
                            curLineItem.PartQty = curLine(FileSchema("PART QTY"))
                            curLineItem.ScheduleNumber = curLine(FileSchema("SCHEDULE NUMBER"))
                            curLineItem.Seq = curLine(FileSchema("SEQ"))
                            curLineItem.PONumber = curLine(FileSchema("PO NUMBER"))
                            curLineItem.ShipTo = curLine(FileSchema("SHIP TO"))
                            curLineItem.POLineNumber = curLine(FileSchema("PO LINE NUMBER"))
                            curLineItem.ShipmentNumber = curLine(FileSchema("SHIPMENT NUMBER"))
                            If curLine(FileSchema("BUCKETIZED FLAG")) = "Y" Then
                                curLineItem.BucketizedFlag = True
                            Else
                                curLineItem.BucketizedFlag = False
                            End If
                        Catch ex As Exception
                            'Something occured with the schema and the PO file is most likely bad.
                            'Exit the sub.
                            Return False
                        End Try

                        '-> Moved the adding of the line item to the dictionary in to the registered product check.

                        'Only execute the following if the product is registered.
                        'We do not want to consider unregistered products when building the buckets and counting totals.

                        If My.Settings.AllowedProductRegistry.Contains(curLineItem.PartNumber) Then
                            'Add the current line item to the dictionary with the key as the sequence number (SEQ)
                            Try
                                LineItems.Add(curLineItem.Seq, curLineItem)
                            Catch ex As Exception
                                'Jump to next line.
                                Continue Do
                            End Try


                            'Create a product entry for the current line item.
                            Dim curProduct As ProductEntry
                            curProduct.PartNumber = curLineItem.PartNumber
                            curProduct.Description = curLineItem.PartDescription
                            curProduct.QtyNeeded = curLineItem.PartQty
                            curProduct.QtyProduced = 0

                            'Need to add the curProduct to the product entries dictionary.
                            If Products.ContainsKey(curProduct.PartNumber) Then
                                'Combine the two products.
                                'Products(curProduct.PartNumber).Combine(curProduct, curProduct.PartNumber)
                                Products(curProduct.PartNumber) = Products(curProduct.PartNumber) + curProduct

                                'Increase the counters.
                                Me.TotalNeeded += curProduct.QtyNeeded
                                Me.TotalProduced += curProduct.QtyProduced
                            Else
                                'Add the product to the product dictionary.
                                Products.Add(curProduct.PartNumber, curProduct)

                                'Increase the counters.
                                Me.TotalNeeded += curProduct.QtyNeeded
                                Me.TotalProduced += curProduct.QtyProduced
                            End If

                            'Need to create a new bucket variable out of the current line item.
                            Dim curBucket As Bucket

                            'Me.Buckets
                            'This bucket is only made up of a single item.
                            Dim curBucketNumber As String = curLineItem.BucketNumber.Split("-")(3)

                            curBucket.BucketNumber = curBucketNumber
                            curBucket.LineItems = New List(Of String)
                            curBucket.LineItems.Add(curLineItem.Seq)
                            curBucket.Products = New Dictionary(Of String, ProductEntry)
                            curBucket.Products.Add(curProduct.PartNumber, curProduct)
                            curBucket.BucketNeeded = curProduct.QtyNeeded
                            curBucket.BucketProduced = 0
                            curBucket.isBulkPack = False

                            'Check if the bucket number is already present in the dictionary.
                            If Buckets.ContainsKey(curBucketNumber) Then
                                'Add the information to the existing bucket.
                                'Buckets(curBucketNumber).Combine(curBucket)
                                Buckets(curBucketNumber) = Buckets(curBucketNumber) + curBucket
                            Else
                                'Add a new bucket to the dictionary.
                                Buckets.Add(curBucketNumber, curBucket)
                            End If
                        End If

                    Loop
                End Using
            End If
        Catch ex As FileNotFoundException
            'File was not found.
            MessageBox.Show("File does not exist.", "No File Found", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex1 As IOException
            'IO Exception. More than likely the file is opened in another process.
            'Try and copy the file into the secondary iSupplier directory and then restart the loading process.

            'Verify the exception is the one we think it is.
            If ex1.HResult = -2147024864 Then
                'Since the error is the one we think it is we can proceed as expected.

                'If the file in the new location exists then we must delete that file so we can copy the file back over in it's updated state.
                If My.Computer.FileSystem.FileExists(Path.Replace(My.Settings.iSupplier_Default, My.Settings.iSupplier_CopiedLocation)) Then
                    My.Computer.FileSystem.DeleteFile(Path.Replace(My.Settings.iSupplier_Default, My.Settings.iSupplier_CopiedLocation))
                End If
                My.Computer.FileSystem.CopyFile(Path, $"{Path.Replace(My.Settings.iSupplier_Default, My.Settings.iSupplier_CopiedLocation)}")

                'Changes the path to match the new location.
                Path = Path.Replace(My.Settings.iSupplier_Default, My.Settings.iSupplier_CopiedLocation)

                If File.Exists(Path) Then

                    'Determines what order of columns the file has bases on the first row.
                    Dim FileSchema As New Dictionary(Of String, Integer)

                    'Start parsing the file.
                    Using inFile As StreamReader = File.OpenText(Path)
                        Dim tempLineVar As List(Of String) = inFile.ReadLine.Split(",").ToList()
                        Dim headerVar = 0

                        FileSchema = IdentifyFileSchema(tempLineVar)

                        'Repeats until the end of the file.
                        Do While Not inFile.EndOfStream AndAlso FileSchema IsNot Nothing
                            Dim curLine As List(Of String) = Nothing
                            If tempLineVar.Count > 1 Then
                                curLine = inFile.ReadLine().Split(",").ToList()
                            Else
                                curLine = inFile.ReadLine().Split(vbTab).ToList()
                            End If
                            'Add statements to check against other forms of deleminators.

                            If headerVar = 0 Then
                                Me.ScheduleNumber = curLine(1).Replace($"-{curLine(FileSchema("SCHEDULE NUMBER")).Split("-")(3)}", "")
                                Me.PONumber = curLine(FileSchema("PO NUMBER"))
                                'Add 1 to headerVar as to not repeat the process a second time.
                                headerVar = 1
                            End If


                            'Me.LineItems
                            'Get the line item and add it to the dictionary.
                            Dim curLineItem As LineItem
                            Try
                                curLineItem.Aliass = curLine(FileSchema("ALIAS"))
                                curLineItem.BaseModel = curLine(FileSchema("BASE MODEL"))
                                curLineItem.BucketNumber = curLine(FileSchema("BUCKET NUMBER"))
                                curLineItem.NeedByDate = curLine(FileSchema("NEED BY DATE"))
                                curLineItem.PartDescription = curLine(FileSchema("PART DESCRIPTION"))
                                curLineItem.PartNumber = curLine(FileSchema("PART NUMBER"))
                                curLineItem.PartQty = curLine(FileSchema("PART QTY"))
                                curLineItem.ScheduleNumber = curLine(FileSchema("SCHEDULE NUMBER"))
                                curLineItem.Seq = curLine(FileSchema("SEQ"))
                                curLineItem.PONumber = curLine(FileSchema("PO NUMBER"))
                                curLineItem.ShipTo = curLine(FileSchema("SHIP TO"))
                                curLineItem.POLineNumber = curLine(FileSchema("PO LINE NUMBER"))
                                curLineItem.ShipmentNumber = curLine(FileSchema("SHIPMENT NUMBER"))
                                If curLine(FileSchema("BUCKETIZED FLAG")) = "Y" Then
                                    curLineItem.BucketizedFlag = True
                                Else
                                    curLineItem.BucketizedFlag = False
                                End If
                            Catch ex As Exception
                                'Something occured with the schema and the PO file is most likely bad.
                                'Exit the sub.
                                Return False
                            End Try

                            '-> Moved the adding of the line item to the dictionary in to the registered product check.

                            'Only execute the following if the product is registered.
                            'We do not want to consider unregistered products when building the buckets and counting totals.

                            If My.Settings.AllowedProductRegistry.Contains(curLineItem.PartNumber) Then
                                'Add the current line item to the dictionary with the key as the sequence number (SEQ)
                                Try
                                    LineItems.Add(curLineItem.Seq, curLineItem)
                                Catch ex As Exception
                                    'Jump to next line.
                                    Continue Do
                                End Try


                                'Create a product entry for the current line item.
                                Dim curProduct As ProductEntry
                                curProduct.PartNumber = curLineItem.PartNumber
                                curProduct.Description = curLineItem.PartDescription
                                curProduct.QtyNeeded = curLineItem.PartQty
                                curProduct.QtyProduced = 0

                                'Need to add the curProduct to the product entries dictionary.
                                If Products.ContainsKey(curProduct.PartNumber) Then
                                    'Combine the two products.
                                    'Products(curProduct.PartNumber).Combine(curProduct, curProduct.PartNumber)
                                    Products(curProduct.PartNumber) = Products(curProduct.PartNumber) + curProduct

                                    'Increase the counters.
                                    Me.TotalNeeded += curProduct.QtyNeeded
                                    Me.TotalProduced += curProduct.QtyProduced
                                Else
                                    'Add the product to the product dictionary.
                                    Products.Add(curProduct.PartNumber, curProduct)

                                    'Increase the counters.
                                    Me.TotalNeeded += curProduct.QtyNeeded
                                    Me.TotalProduced += curProduct.QtyProduced
                                End If

                                'Need to create a new bucket variable out of the current line item.
                                Dim curBucket As Bucket

                                'Me.Buckets
                                Dim curBucketNumber As String = curLineItem.BucketNumber.Split("-")(3)

                                curBucket.BucketNumber = curBucketNumber
                                curBucket.LineItems = New List(Of String)
                                curBucket.LineItems.Add(curLineItem.Seq)
                                curBucket.Products = New Dictionary(Of String, ProductEntry)
                                curBucket.Products.Add(curProduct.PartNumber, curProduct)
                                curBucket.BucketNeeded = curProduct.QtyNeeded
                                curBucket.BucketProduced = 0
                                curBucket.isBulkPack = False

                                'Check if the bucket number is already present in the dictionary.
                                If Buckets.ContainsKey(curBucketNumber) Then
                                    'Add the information to the existing bucket.
                                    'Buckets(curBucketNumber).Combine(curBucket)
                                    Buckets(curBucketNumber) = Buckets(curBucketNumber) + curBucket
                                Else
                                    'Add a new bucket to the dictionary.
                                    Buckets.Add(curBucketNumber, curBucket)
                                End If
                            End If

                        Loop
                    End Using
                End If
            End If
        End Try

        Return True
    End Function

    'Allows users to save the PO in it's updated state.
    Public Sub Save()

        If Buckets IsNot Nothing AndAlso Buckets.Count > 0 Then
            Dim json As String = JsonConvert.SerializeObject(Me, Formatting.Indented)

            'Writes data to the PO save location.
            System.IO.File.WriteAllText($"{My.Settings.POSaveDirectory}{Me.PONumber}.JSON", json)
        End If
    End Sub

    Public Function Load(PONumber As String)
        Dim FileName = $"{My.Settings.POSaveDirectory}{PONumber}.JSON"
        If File.Exists(FileName) Then
            Dim Json As String = File.ReadAllText(FileName)

            Dim tempPO As ProductOrder = JsonConvert.DeserializeObject(Of ProductOrder)(Json)

            Me.PONumber = PONumber
            Me.OriginFile = tempPO.OriginFile
            Me.ScheduleNumber = tempPO.ScheduleNumber
            Me.TotalNeeded = tempPO.TotalNeeded
            Me.TotalProduced = tempPO.TotalProduced

            Me.Buckets = tempPO.Buckets
            Me.LineItems = tempPO.LineItems
            Me.Products = tempPO.Products
            Me.Inputs = tempPO.Inputs

            Me.RogueBuckets = tempPO.RogueBuckets
            Me.BulkPacks = tempPO.BulkPacks
        End If
    End Function

    'Resets the counts and progress indicators of the PO to the default state.
    Public Sub Reset()
        'Loop through each bucket variable for the PO and reset the counts.
        For Each bucketVar As Bucket In Buckets.Values
            'Reset the produced counts.
            bucketVar.BucketProduced = 0

            'Loop through individual products.
            For Each productVar As ProductEntry In bucketVar.Products.Values
                'Reset the produced counts.
                productVar.QtyProduced = 0
            Next

            'At this point all production recordings for the specified bucket have been reset to zero.
        Next

        'Loop through the products dictionary and reset the produced counts.
        For Each productVar As ProductEntry In Products.Values
            'Reset produced count.
            productVar.QtyProduced = 0
        Next

        'Reset the global PO counters.
        TotalProduced = 0

        'PO should be completely reset to it's default state.

        'Does not save in case the user didn't intend for it to be called.
        'PO must be saved manually at this point by the user.
    End Sub

    'Resets a specified bucket production history.
    'Overload 1: Input is a Bucket.
    Public Sub ResetBucket(Bucket As Bucket)
        'Reset all production counts in the bucket.

        'Reset all production counts for the products in both the bucket and product dictionary.

    End Sub

    'Overload 2: Input is a bucket reference number.
    Public Sub ResetBucket(BucketNumber As String)
        'Reset all production counts in the bucket.

        'Reset all production counts for the products in both the bucket and product dictionary.

    End Sub

    Private Function IdentifyFileSchema(headerRow As List(Of String)) As Dictionary(Of String, Integer)
        Dim tempSchema As New Dictionary(Of String, Integer)

        'Loop through header row and determine the indexes of each header.
        For Each index As String In headerRow
            If Not tempSchema.ContainsKey(index) Then
                tempSchema.Add(index, headerRow.IndexOf(index))
            End If
        Next

        'Ensure the needed entries are in the schema.
        If tempSchema.Keys.Count > 0 Then
            Return tempSchema
        Else
            Return Nothing
        End If
    End Function

    'Should be either a PO Number or a Schedule Number
    Public Enum Identifier As Integer
        PONumber = 0
        ScheduleNumber = 1
    End Enum

    'Identifier is either a PO Number or a Schedule Number.
    Private Shared Function CheckSignature(Identifier As String) As Identifier
        'PO Numbers will only contain 1 "-".
        'Schedule Numbers will contain at least 2 "-".

        'Check how many "-" are present by splitting the string and counting number of children.
        Dim NumberOfDashes As Integer = Identifier.Split("-").Count - 1

        If NumberOfDashes = 1 Then
            'PO Number.
            Dim isPONumber As Boolean = True

            'Check if the first item has 4 or 5 characters.
            If Identifier.Split("-")(0).Count <> 4 And Identifier.Split("-")(0).Count <> 5 Then
                isPONumber = False
            End If

            'Check if the second item has 5 characters.
            If Identifier.Split("-")(1).Count <> 5 Then
                isPONumber = False
            End If

            If isPONumber = True Then
                Return ProductOrder.Identifier.PONumber
            End If

        ElseIf NumberOfDashes >= 3 Then
            'Schedule Number.
            Dim isScheduleNumber As Boolean = True

            'Check if location of the date is the correct size.
            If Identifier.Split("-")(2).Count <> 6 Then
                isScheduleNumber = False
            End If

            'Check if the ship-to identifier is made up only by letters.
            If IsNumeric(Identifier.Split("-")(0)) = True Then
                isScheduleNumber = False
            End If

            'Check to ensure the department number is made up only by numbers.
            If isOnlyNumeric(Identifier.Split("-")(1)) = False Then
                isScheduleNumber = False
            End If

            If isScheduleNumber = True Then
                Return ProductOrder.Identifier.ScheduleNumber
            End If
        End If

        Return Nothing
    End Function

    Public Shared Function FindPO(PONumber As String) As ProductOrder
        'Loop through the files in the specified directory in order to find the identified string.

        If CheckSignature(PONumber) = ProductOrder.Identifier.PONumber Then
            'Declare return value.
            Dim returnPO As ProductOrder = Nothing

            'Check the PO save location for the PO JSON file.
            If File.Exists($"{My.Settings.POSaveDirectory}{PONumber}.JSON") Then

                returnPO.Load(PONumber)

                Return returnPO
            End If

            'Check the files in the iSupplier folder.
            If File.Exists($"{My.Settings.iSupplier_Default}{PONumber}.txt") Then

                returnPO = New ProductOrder(PONumber)

                Return returnPO
            End If

        End If

        'No PO was found. Return nothing.
        Return Nothing
    End Function

    'Schedule Number is only the date at the end of the schedule number.
    Public Shared Function FindPOs(ScheduleNumber As String) As List(Of ProductOrder)
        'Loop through the files in the specified directory in order to find the identified string.

        If CheckSignature(ScheduleNumber) = ProductOrder.Identifier.ScheduleNumber Then
            'Declare return value.
            Dim returnList As New List(Of ProductOrder)

            Dim searchString As String = $"-{ScheduleNumber.Split("-").Last()}"
            Dim files As List(Of String) = My.Computer.FileSystem.FindInFiles(My.Settings.iSupplier_Default, searchString, True, FileIO.SearchOption.SearchTopLevelOnly).ToList()

            For Each fileVar In files
                'Add the PO to the return list after having created a new PO object.
                returnList.Add(New ProductOrder(fileVar.Split("\").Last().Replace(".txt", "")))
            Next

            Return returnList
        End If

        'No PO's were found. Return nothing.
        Return Nothing
    End Function

    Public Function BucketizeProducts() As List(Of BucketizedProductEntry)
        Dim returnList As New List(Of BucketizedProductEntry)

        'Loop through each product in each of the buckets.
        For Each bucket In Buckets.Values

            'Loop through products.
            For Each productVar In bucket.Products.Values

                'Add the bucketized product to the list.
                returnList.Add(productVar.Bucketize(bucket.BucketNumber))

            Next
        Next

        'Return the variable.
        Return returnList
    End Function

    Public Sub Produce(PartNumber As String, QtyProduced As Double)

        'Check that the part number is valid.
        If Me.Products.ContainsKey(PartNumber) Then
            'Need to update that count in all locations that the product exists in.

            'ProductEntry:

            Me.Products(PartNumber).QtyProduced += QtyProduced

            'Buckets:

            Dim QtyExtra As Double = 0
            Dim index As Integer = 0
            While index < Buckets.Count

                Dim bucketVar As Bucket = Buckets.ElementAt(index).Value

                'Assume the buckets are in order from first to last.
                If bucketVar.Products.ContainsKey(PartNumber) Then

                    'Check if the QtyProduced will exceed the needed quantity.
                    If bucketVar.Products(PartNumber).QtyProduced + QtyProduced >= bucketVar.Products(PartNumber).QtyNeeded Then
                        'Fill the bucket and then reset the QtyExtra variable to reflect the proper variable.

                        QtyExtra = QtyExtra - bucketVar.Products(PartNumber).QtyNeeded - bucketVar.Products(PartNumber).QtyProduced
                        bucketVar.Products(PartNumber).QtyProduced = bucketVar.Products(PartNumber).QtyNeeded
                        bucketVar.BucketProduced = bucketVar.BucketNeeded
                    Else
                        'Just add the total to the bucket.
                        bucketVar.Products(PartNumber).QtyProduced += QtyProduced
                        bucketVar.BucketProduced += QtyProduced
                    End If
                End If

                index += 1
            End While

            'PO Counts:
            Me.TotalProduced += QtyProduced
        End If

        Me.Save()
    End Sub
End Class
