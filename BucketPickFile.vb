Imports System.IO
Imports Product_Breakdown.Application

Public Class BucketPickFile
    'Each morning the team lead needs to initiate the schedule pull by refreshing the live feed.
    'This will most likely be a button on the main input page.

    'File path to the schedule information will be set upon the start of the application if the user has not started the application on the specified device before.
    'File path should be in the same root dirctory as the original file is pulled from.

    Public Property BucketNumber As String  '0
    Public Property ScheduleNumber As String    '1
    Public Property AliasVal As String  '2
    Public Property BaseModel As String '3
    Public Property PartNumber As String    '4
    Public Property PartDescription As String   '5
    Public Property PartQuantity As Integer '6
    Public Property Sequence As Integer '7
    Public Property PONumber As String  '8
    Public Property ShipToLocation As String    '9
    Public Property POLineNumber As Integer '10
    Public Property ShipmentNumber As Integer   '11
    Public Property NeedByDate As Date  '12
    Public Property BucketizedFlag As Boolean   '13     Y = True, N = False

    'Need to generate a dictionary(of tuples) for each file as that is what will be displayed on the main window.
    Public Property ProductionNumbers As Dictionary(Of String, List(Of Integer)) '(PartNumber, (Needed, Produced))
    Private Property POOrderTotal As Integer 'Stores the total number of ordered parts for the PO.
    Public Property OriginFile As String

    'We assume that the file selected is only for Terrace Frames and nothing else.

    Public Sub New(PONumber As String)
        Me.ProductionNumbers = New Dictionary(Of String, List(Of Integer))

        'Made to be ambiguous.
        LoadFile(PONumber)
        Me.OriginFile = PONumber
    End Sub

    Public Sub LoadFile(Path As String)
        'Create new string for the path of the PO from the saved general path in app.settings and the PONumber variable.
        Try
            If File.Exists(Path) Then
                'We want the summation of all the product orders in the given file not individual orders for each.

                'Start parsing the file.
                Using inFile As StreamReader = File.OpenText(Path)
                    Dim tempLineVar As List(Of String) = inFile.ReadLine.Split(",").ToList()
                    Dim headerVar = 0
                    'Repeats until the end of the file.
                    Do While Not inFile.EndOfStream
                        Dim curLine As List(Of String) = Nothing
                        If tempLineVar.Count > 1 Then
                            curLine = inFile.ReadLine().Split(",").ToList()
                        Else
                            curLine = inFile.ReadLine().Split(vbTab).ToList()
                        End If
                        'Add statements to check against other forms of deleminators.

                        If headerVar = 0 Then
                            Me.ScheduleNumber = curLine(1).Replace($"-{curLine(1).Split("-")(3)}", "")
                            Me.PONumber = curLine(8)
                            'Add 1 to headerVar as to not repeat the process a second time.
                            headerVar = 1
                        End If

                        'Should only bring in products that match the part numbers for Terrace & Stride Frames.
                        If ProductDirectory.Keys.Contains(curLine(4)) Then
                            If ProductionNumbers.Keys.Contains(curLine(4)) Then
                                'Gets the current number of parts stored in the dictionary.
                                Dim listVar = ProductionNumbers(curLine(4))
                                'Sums the new total from the current total and the new amount from the line.
                                listVar(0) = listVar(0) + CInt(curLine(6))
                                'Sets the newest total number of parts needed in the ProductionNumbers for the main window.
                                ProductionNumbers(curLine(4)) = listVar
                                'Adds the new parts to the order total.
                                Me.POOrderTotal += CInt(curLine(6))
                            ElseIf Not ProductionNumbers.Keys.Contains(curLine(4)) Then
                                'Add the new part number to the ProductionNumbers dictionary.
                                ProductionNumbers.Add(curLine(4), New List(Of Integer) From {CInt(curLine(6)), 0})
                                'Adds the new parts to the order total.
                                Me.POOrderTotal += CInt(curLine(6))
                            End If
                        End If
                    Loop
                End Using
            End If
        Catch ex As FileNotFoundException
            'File was not found.
            MessageBox.Show("File does not exist.", "No File Found", MessageBoxButton.OK, MessageBoxImage.Warning)
        End Try
    End Sub
End Class
