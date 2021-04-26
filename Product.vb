Imports Newtonsoft.Json
Imports System.IO
Imports Product_Breakdown.Application

Public Class Product
    Public Property PartNumber As String
    Public Property Description As String
    Public Property ProductionArea As String
    Public Property ImageKey As Integer
    Public Property Dimensions As New List(Of Measurement)
    Public Property Path As String
    Public Property Measurements As New Dictionary(Of String, Measurement)
    Public Property MeasurementGroups As New Dictionary(Of String, List(Of String))
    Public Property PrintPath As String
    Public Property Fixtured As Boolean 'distinguishes whether the process of creating the product requires a fixture or not. Value is true if a fixture is used for production.

    'The first-piece and audit properties do not need to be filled.
    'Primarily for future changes.
    Public Property FirstPiece_QS As String 'ID for the quality sheet for first-piece inspections.
    Public Property Audit_QS As String 'ID for the quality sheet for the audit inspections.

    'Destinguishes the part from others apart from the part number.
    Public Property Family As String 'Specifies the Family the product belongs to. Example: Terrace.
    Public Property SubFamily As String 'Specifies the sub-family the product belongs to. Example: Frame or Vertical.
    Public Property Classification As String    'Specifies the product's classification. Example: Base Height or Stacker.

    Public Property UI_ButtonDescription As String 'Text that appears on the button that navigates the user to the appropriate quality sheet.

    Public ProductAssemblyInformation As AssemblyInformation    'Holds information about how the product is assembled from child products if applicable. If not then only the ChildProduct1 variable is filled with the parent product number.

    <JsonIgnore> Private Shared FileName As String = Nothing

    <JsonConstructor>
    Public Sub New(Optional PartNumber As String = Nothing, Optional Description As String = Nothing, Optional ProductionArea As String = Nothing, Optional ImageKey As Integer = 0, Optional Dimensions As List(Of Measurement) = Nothing, Optional Path As String = Nothing, Optional Measurements As Dictionary(Of String, Measurement) = Nothing, Optional MeasurementGroups As Dictionary(Of String, List(Of String)) = Nothing, Optional PrintPath As String = "", Optional Fixtured As Boolean = False, Optional FirstPiece As String = Nothing, Optional Audit As String = Nothing, Optional ButtonDescription As String = Nothing, Optional Family As String = Nothing, Optional SubFamily As String = Nothing, Optional Classification As String = Nothing, Optional assemblyInfo As AssemblyInformation = Nothing)
        Me.PartNumber = PartNumber
        Me.Description = Description
        Me.ProductionArea = ProductionArea
        Me.ImageKey = ImageKey
        If Dimensions Is Nothing Then
            Me.Dimensions = New List(Of Measurement)
        Else
            Me.Dimensions = Dimensions
        End If
        Me.Path = Path
        If Measurements Is Nothing Then
            Me.Measurements = New Dictionary(Of String, Measurement)
        Else
            Me.Measurements = Measurements
        End If
        If MeasurementGroups Is Nothing Then
            Me.MeasurementGroups = New Dictionary(Of String, List(Of String))
        Else
            Me.MeasurementGroups = MeasurementGroups
        End If
        Me.PrintPath = PrintPath
        Me.Fixtured = Fixtured
        Me.FirstPiece_QS = FirstPiece
        Me.Audit_QS = Audit
        Me.UI_ButtonDescription = ButtonDescription
        Me.Family = Family
        Me.SubFamily = SubFamily
        Me.Classification = Classification
        Me.ProductAssemblyInformation = assemblyInfo
    End Sub
    'Functions for use with the Product Class

    'Loads information relevant to the part number passed in.
    'Input // None.
    'Output // None.
    'Result // Generates a product.
    Public Shared Function LoadProduct(PartNumber As String) As Product
        FileName = $"{My.Settings.ProductPath}\{PartNumber}.json"
        If File.Exists(FileName) Then
            Dim Json As String = File.ReadAllText(FileName)
            Return JsonConvert.DeserializeObject(Of Product)(Json)
        End If
    End Function

    'Send information from ProductDirectory to the Product Variables file after a product was added or deleted
    'Input // None.
    'Output // None.
    'Result // Adds a product To the disctionary and appends line to the product variable list file stored on the server.
    Public Sub SaveProduct()
        Try
            ProductDirectory.Add(Me.PartNumber, Me)
        Catch ex As Exception
        End Try
        Dim FileName As String = $"{My.Settings.ProductPath}\{Me.PartNumber}.json"
        Dim Json As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        File.WriteAllText(FileName, Json)
    End Sub

    'Deletes specified product by overwriting the file that housed the old list of parts with a new list generated programmatically that drops the chosen part
    'Input // None.
    'Output // None.
    'Result // Removes a specific product from the product list.
    Public Shared Sub RemoveProductList(oldProduct As Product)
        Try
            ProductDirectory.Remove(oldProduct.PartNumber)
        Catch ex As Exception
        End Try
        FileName = $"{oldProduct.Path}\{oldProduct.PartNumber}"
        File.Delete(FileName)
    End Sub

    Public Shared Sub GenerateProductList()
        For Each ProductFile In FileIO.FileSystem.GetFiles(My.Settings.ProductPath)
            Dim tempProduct As Product = LoadProduct(ProductFile.Split("\").Reverse.ElementAt(0).Split(".").ElementAt(0))
            Try
                ProductDirectory.Add(tempProduct.PartNumber, tempProduct)
            Catch ex As Exception
                'Impliment error logging maybe.
            End Try
        Next
    End Sub

    Structure AssemblyInformation
        'Lists the parts required for completion of the production.
        'If the product is a frame it will list all child products needed to complete 1 frame.
        'If the product is anything that doesn't require assembly or more than one product to complete then it will list just the part number of the product itself.

        'Tuple containing [Part Number,Qty Needed]

        'ChildProduct1 will always be filled.
        Public ChildProduct1 As Tuple(Of String, Integer)

        Public ChildProduct2 As Tuple(Of String, Integer)

        Public ChildProduct3 As Tuple(Of String, Integer)

        Public ChildProduct4 As Tuple(Of String, Integer)

        Public ChildProduct5 As Tuple(Of String, Integer)

        Public ChildProduct6 As Tuple(Of String, Integer)

        Public ChildProduct7 As Tuple(Of String, Integer)

        Public ChildProduct8 As Tuple(Of String, Integer)

        Public ChildProduct9 As Tuple(Of String, Integer)

        Public ChildProduct10 As Tuple(Of String, Integer)

        'Lists how many product numbers are present.
        Public ChildProductsCount As Integer

    End Structure
End Class

Public Class Measurement
    'Holds all information required to auto generate a measurement check control in a quality checksheet.

    'Properties Include:
    '- Product Host: String that describes the Product Number associated with the measurement.
    '- ID: String that describes the unique identification number of the measurement.
    '- Name: String that describes the name of the measurement. As dispalyed by forms that use the measurement.
    '- Type: String value that describes the type of measurement.
    '- Max: Decimal Value describing the maximum tolerance.
    '- Nom: Decimal Value describing the nominal tolerance.
    '- Min: Decimal Value describing the minimum tolerance.
    '- IndexInQS: Integer Value that holds the location information for where the dimensions are located for in the product file.
    '- DisplayOrder: Integer value describing what the index of the measurement should be in the history recall form.
    '- LastUpdated: String value describing the last time the measurement was changed.
    '- UpdatedBy: Integer value of the User ID Number of the user that last updated the measurement.

    'Maybe add a shortcut key to allow users quick access to it

    '---Properties---
    Public Property ID As String
    Public Property Name As String
    Public Property Level As String
    Public Property Type As String
    Public Property Maximum As Decimal
    Public Property Nominal As Decimal
    Public Property Minimum As Decimal
    Public Property DisplayOrder As Integer
    Public Property LastUpdated As DateTimeOffset
    Public Property UpdatedBy As Integer

    <JsonIgnore> Private Shared FileName As String = Nothing

    <JsonConstructor>
    Public Sub New(Optional ID As String = Nothing, Optional Name As String = Nothing, Optional Level As String = Nothing, Optional Type As String = Nothing, Optional Maximum As Decimal = Nothing, Optional Nominal As Decimal = Nothing, Optional Minimum As Decimal = Nothing, Optional DisplayOrder As Integer = Nothing, Optional LastUpdated As DateTimeOffset = Nothing, Optional UpdatedBy As Integer = Nothing)
        Me.ID = ID
        Me.Name = Name
        Me.Level = Level
        Me.Type = Type
        Me.Maximum = Maximum
        Me.Nominal = Nominal
        Me.Minimum = Minimum
        Me.DisplayOrder = DisplayOrder
        If LastUpdated = Nothing Then
            Me.LastUpdated = DateTimeOffset.Now
        Else
            Me.LastUpdated = LastUpdated
        End If
        Me.UpdatedBy = UpdatedBy
    End Sub

    Public Shared Function LoadMeasurement(ID As String) As Measurement
        FileName = $"{"P:\Quality_Wizard\Quality_Wizard_Data_Files\Product Info\"}{ID}.json"
        If File.Exists(FileName) Then
            Dim Json As String = File.ReadAllText(FileName)
            Return JsonConvert.DeserializeObject(Of Measurement)(Json)
        End If
    End Function

    Public Function SaveMeasurement()
        'MeasurementList.Add(Me.ID, Me)
        FileName = $"{"P:\Quality_Wizard\Quality_Wizard_Data_Files\Product Info\"}{Me.ID}.json"
        Dim Json As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        File.WriteAllText(FileName, Json)
    End Function
End Class