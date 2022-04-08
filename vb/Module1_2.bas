Attribute VB_Name = "Module1"

Public sql As String
Public c As ADODB.Connection
Public r As ADODB.Recordset


Function Connection()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;Password=lnt123;User ID=LNT;Persist Security Info=True"
Set r = New ADODB.Recordset
End Function
