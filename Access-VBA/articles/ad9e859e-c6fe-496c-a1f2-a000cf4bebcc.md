
# QueryDef.Execute Method (DAO)

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)
[About the Contributors](#AboutContributors)


Executes an SQL statement on the specified object.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Execute**( ** _Options_** )

 _expression_ A variable that represents a **QueryDef** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Options_|Optional|**Variant**||

## Remarks
<a name="sectionSection1"> </a>

You can use the following  **[RecordsetOptionEnum](3a9d8664-dcb6-cb60-7cf6-e229eb699ef1.md)** constants for _Options_.



|**Constant**|**Description**|
|:-----|:-----|
|**dbDenyWrite**|Denies write permission to other users (Microsoft Access workspaces only).|
|**dbInconsistent**|(Default) Executes inconsistent updates (Microsoft Access workspaces only).|
|**dbConsistent**|Executes consistent updates (Microsoft Access workspaces only).|
|**dbSQLPassThrough**|Executes an SQL pass-through query. Setting this option passes the SQL statement to an ODBC database for processing (Microsoft Access workspaces only).|
|**dbFailOnError**|Rolls back updates if an error occurs (Microsoft Access workspaces only).|
|**dbSeeChanges**|Generates a run-time error if another user is changing data you are editing (Microsoft Access workspaces only).|
|**dbRunAsync**|Executes the query asynchronously (ODBCDirect Connection and  **QueryDef** objects only).|
|**dbExecDirect**|Executes the statement without first calling  **SQLPrepare** ODBC API function (ODBCDirect Connection and **QueryDef** objects only).|

 **Note**  ODBCDirect workspaces are not supported in Microsoft Access 2013. Use ADO if you want to access external data sources without using the Microsoft Access database engine.




 **Note**  The constants  **dbConsistent** and **dbInconsistent** are mutually exclusive. You can use one or the other, but not both in a given instance of **OpenRecordset**. Using both **dbConsistent** and **dbInconsistent** causes an error.

Use the  **[RecordsAffected](29a864b5-305c-d33f-b2ca-fc9a08baaa5c.md)** property of the **[Connection](f469b04e-2539-6b53-31f2-85fe22fcc2fc.md)**, **[Database](6cf2ddf8-3957-a15e-5eeb-85f81c1e415e.md)**, or **[QueryDef](0b3d901c-345d-42a2-f5f1-fb09cc562e27.md)** object to determine the number of records affected by the most recent **[Execute](ad9e859e-c6fe-496c-a1f2-a000cf4bebcc.md)** method. For example, **RecordsAffected** contains the number of records deleted, updated, or inserted when executing an action query. When you use the **Execute** method to run a query, the **RecordsAffected** property of the **QueryDef** object is set to the number of records affected.

In a Microsoft Access workspace, if you provide a syntactically correct SQL statement and have the appropriate permissions, the  **Execute** method won't fail — even if not a single row can be modified or deleted. Therefore, always use the **dbFailOnError** option when using the **Execute** method to run an update or delete query. This option generates a run-time error and rolls back all successful changes if any of the records affected are locked and can't be updated or deleted.

In earlier versions of the Microsoft Jet database engine, SQL statements were automatically embedded in implicit transactions. If part of a statement executed with  **dbFailOnError** failed, the entire statement would be rolled back. To improve performance, these implicit transactions were removed starting with version 3.5. If you are updating older DAO code, be sure to consider using explicit transactions around **Execute** statements.

For best performance in a Microsoft Access workspace, especially in a multiuser environment, nest the  **Execute** method inside a transaction. Use the **[BeginTrans](aa7c3bf8-fb08-9360-5998-4bf3f721ecbb.md)** method on the current **[Workspace](bf3ab863-5e9a-4842-1f82-2ccf958d9779.md)** object, then use the **Execute** method, and complete the transaction by using the **[CommitTrans](e6d129fb-a578-5c79-9c16-6444519f0daf.md)** method on the **Workspace**. This saves changes on disk and frees any locks placed while the query is running.


## Example
<a name="sectionSection2"> </a>

This example demonstrates the  **Execute** method when run from both a **QueryDef** object and a **Database** object. The **ExecuteQueryDef** and **PrintOutput** procedures are required for this procedure to run.


```vb
Sub ExecuteX() 
 
   Dim dbsNorthwind As Database 
   Dim strSQLChange As String 
   Dim strSQLRestore As String 
   Dim qdfChange As QueryDef 
   Dim rstEmployees As Recordset 
   Dim errLoop As Error 
 
   ' Define two SQL statements for action queries. 
   strSQLChange = "UPDATE Employees SET Country = " &; _ 
      "'United States' WHERE Country = 'USA'" 
   strSQLRestore = "UPDATE Employees SET Country = " &; _ 
      "'USA' WHERE Country = 'United States'" 
 
   Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
   ' Create temporary QueryDef object. 
   Set qdfChange = dbsNorthwind.CreateQueryDef("", _ 
      strSQLChange) 
   Set rstEmployees = dbsNorthwind.OpenRecordset( _ 
      "SELECT LastName, Country FROM Employees", _ 
      dbOpenForwardOnly) 
 
   ' Print report of original data. 
   Debug.Print _ 
      "Data in Employees table before executing the query" 
   PrintOutput rstEmployees 
    
   ' Run temporary QueryDef. 
   ExecuteQueryDef qdfChange, rstEmployees 
    
   ' Print report of new data. 
   Debug.Print _ 
      "Data in Employees table after executing the query" 
   PrintOutput rstEmployees 
 
   ' Run action query to restore data. Trap for errors, 
   ' checking the Errors collection if necessary. 
   On Error GoTo Err_Execute 
   dbsNorthwind.Execute strSQLRestore, dbFailOnError 
   On Error GoTo 0 
 
   ' Retrieve the current data by requerying the recordset. 
   rstEmployees.Requery 
 
   ' Print report of restored data. 
   Debug.Print "Data after executing the query " &; _ 
      "to restore the original information" 
   PrintOutput rstEmployees 
 
   rstEmployees.Close 
    
   Exit Sub 
    
Err_Execute: 
 
   ' Notify user of any errors that result from 
   ' executing the query. 
   If DBEngine.Errors.Count > 0 Then 
      For Each errLoop In DBEngine.Errors 
         MsgBox "Error number: " &; errLoop.Number &; vbCr &; _ 
            errLoop.Description 
      Next errLoop 
   End If 
    
   Resume Next 
 
End Sub 
 
Sub ExecuteQueryDef(qdfTemp As QueryDef, _ 
   rstTemp As Recordset) 
 
   Dim errLoop As Error 
    
   ' Run the specified QueryDef object. Trap for errors, 
   ' checking the Errors collection if necessary. 
   On Error GoTo Err_Execute 
   qdfTemp.Execute dbFailOnError 
   On Error GoTo 0 
 
   ' Retrieve the current data by requerying the recordset. 
   rstTemp.Requery 
    
   Exit Sub 
 
Err_Execute: 
 
   ' Notify user of any errors that result from 
   ' executing the query. 
   If DBEngine.Errors.Count > 0 Then 
      For Each errLoop In DBEngine.Errors 
         MsgBox "Error number: " &; errLoop.Number &; vbCr &; _ 
            errLoop.Description 
      Next errLoop 
   End If 
    
   Resume Next 
 
End Sub 
 
Sub PrintOutput(rstTemp As Recordset) 
 
   ' Enumerate Recordset. 
   Do While Not rstTemp.EOF 
      Debug.Print "  " &; rstTemp!LastName &; _ 
         ", " &; rstTemp!Country 
      rstTemp.MoveNext 
   Loop 
 
End Sub
```

The following example shows how to execute a parameter query. The  **Parameters** collection is used to set the **Organization** parameter of the **myActionQuery** query before the query is executed.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.mdl) |[About the Contributors](#AboutContributors)




```vb
Public Sub ExecParameterQuery()

    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef

    Set dbs = CurrentDb
    Set qdf = dbs.QueryDefs("myActionQuery")

    'Set the value of the QueryDef's parameter
    qdf.Parameters("Organization").Value = "Microsoft"

    'Execute the query
    qdf.Execute dbFailOnError

    'Clean up
    qdf.Close
    Set qdf = Nothing
    Set dbs = Nothing

End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 

