### Refresh One Query

If you want to just refresh one query and wait till it finished then the sample code would be

```
    WorksheetCodeName.ListObjects("TableName").QueryTable.Refresh False   
    'This last false will say that wait till it is refreshed.   
    'Sample Example  
    Sheet1.ListObjects("Table_Get_Data_From_JSON_File").QueryTable.Refresh False
```

### Refresh 2-3 Query in sequence

We can follow the same approach like above here. I am just going to add sample code

```
    RelevantDataSheet.ListObjects("KeepRelevantData").QueryTable.Refresh False   
    DoEvents  
    AllCombinationSheet.ListObjects("AllCombination").QueryTable.Refresh False  
    DoEvents  
    RelevantDataSheet.ListObjects("DateRange").QueryTable.Refresh False  
    DoEvents
```

### Refresh All Query

There is not a direct approach to wait before all query refreshed. So we have to use loop here. Here is a sample code for that.

```
    Dim CurrentSheet As Worksheet  
    Dim CurrentQueryTable As QueryTable  
    For Each CurrentSheet In ApplyInWorkbook.Worksheets  
        For Each CurrentQueryTable In CurrentSheet.QueryTables  
            CurrentQueryTable.BackgroundQuery = False  
        Next CurrentQueryTable  
    Next CurrentSheet  
```

### A utility sub to Refresh Query

```
Private Enum QueryType
    OLEDB = 1
    ODBC = 2
End Enum

Private Sub WaitToRefreshQuery(QueryName As String, TypeOfQuery As QueryType)
    
    Dim DefaultRefresh As Boolean
    Dim Connection As WorkbookConnection
    
    Set Connection = ThisWorkbook.Connections(QueryName)
    Dim ConnectionType As Object
    If TypeOfQuery = ODBC Then
        Set ConnectionType = Connection.ODBCConnection
    ElseIf TypeOfQuery = OLEDB Then
        Set ConnectionType = Connection.OLEDBConnection
    End If
    With ConnectionType
        DefaultRefresh = .BackgroundQuery
        .BackgroundQuery = False
        .Refresh
        .BackgroundQuery = DefaultRefresh
    End With
    
End Sub
```
