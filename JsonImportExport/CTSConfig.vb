Public Class CTSConfig
    Private Property Connection As String
    Public Sub New(pConn As String)
        Connection = pConn
    End Sub
    Public Sub SetConnect(pConn As String)
        Connection = pConn
    End Sub
    Public Property ConfigKey As String
    Public Property ConfigCode As String
    Public Property ConfigValue As String
    Public Function SaveData() As String
        Return "Save " & ConfigCode & " Complete"
    End Function
End Class
