Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

    Public Shared CurrentPO As BucketPickFile
    Public Shared ProductDirectory As New Dictionary(Of String, Product)

    Public Shared Schema As New Dictionary(Of String, String)

    Public Shared ProductFamily As String
    Public Shared ProductSubFamily As String


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
End Class
