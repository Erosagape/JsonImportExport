Imports System.Reflection
Imports Newtonsoft.Json
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim data = "[{""ConfigCode"":""Code1"",""ConfigKey"":""Key1"",""ConfigValue"":""Value1""},{""ConfigCode"":""Code2"",""ConfigKey"":""Key2"",""ConfigValue"":""Value2""}]"
        Dim json As String = "{""ClassName"":""CTSConfig"",""Database"":""DEVTEST"",""Rows"":" & data & "}"

        Dim obj As JsonImportExport.CDataBox = JsonConvert.DeserializeObject(Of JsonImportExport.CDataBox)(json)
        Dim className As String = obj.ClassName
        Dim dataBase As String = obj.Database
        For Each o As Object In obj.Rows
            Dim str As String = JsonConvert.SerializeObject(o)
            Dim row As JsonImportExport.CTSConfig = JsonConvert.DeserializeObject(Of JsonImportExport.CTSConfig)(str)
            row.SetConnect(dataBase)
            MessageBox.Show(row.SaveData())
        Next
    End Sub
End Class
