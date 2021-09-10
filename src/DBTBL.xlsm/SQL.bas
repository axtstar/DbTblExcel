Attribute VB_Name = "SQL"
Function GetAllTablesString(tables As String) As String
    GetAllTablesString = Sheets("devSheet").Cells(8, 3).Value
    If tables <> "" Then GetAllTablesString = GetAllTablesString & " WHERE TableName like '%" & tables & "%'"
End Function

