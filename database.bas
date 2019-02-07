Attribute VB_Name = "database"
Public cn As New ADODB.Connection

Public Sub Connection()
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=database.mdb;Persist Security Info=False"
End Sub


