---
title: Modify a Query from a DAO Recordset
ms.prod: access
ms.assetid: b5679ca8-9bcd-2d28-15af-2640db727dd4
ms.date: 06/08/2017
---


# Modify a Query from a DAO Recordset

You can use the  **[Requery](http://msdn.microsoft.com/library/A5D66EB5-499C-4133-F6C3-C7A1619A8A11%28Office.15%29.aspx)** method on a dynaset-type or snapshot-type **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object when you want to run the underlying query again after changing a parameter. This is more convenient than opening a new **Recordset**, and it runs faster.

The following code example creates a  **Recordset** object and passes it to a function that uses the **[CopyQueryDef](http://msdn.microsoft.com/library/FEE8C2FE-500E-DFB3-21CE-211E54FF334B%28Office.15%29.aspx)** method to extract the equivalent SQL string. It then prompts the user to add an additional constraint clause to the query. The code uses the **Requery** method to run the modified query.



```vb
Sub AddQuery() 
 
Dim dbsNorthwind As DAO.Database 
Dim qdfSalesReps As DAO.QueryDef 
Dim rstSalesReps As DAO.Recordset 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
 
   Set qdfSalesReps = dbsNorthwind.CreateQueryDef("SalesRepQuery") 
   qdfSalesReps.SQL = "SELECT * FROM Employees WHERE Title = " &; _ 
                      "'Sales Representative'" 
 
   Set rstSalesReps = qdfSalesReps.OpenRecordset() 
 
   ' Call the function to add a constraint. 
   AddQueryFilter rstSalesReps 
 
   ' Return database to original. 
   dbsNorthwind.QueryDefs.Delete "SalesRepQuery" 
 
   rstSalesReps.Close 
   qdfSalesReps.Close 
   dbsNorthwind.Close 
 
   Set rstSalesReps = Nothing 
   Set qdfSalesReps = Nothing 
   Set dbsNorthwind = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub 
 
Sub AddQueryFilter(rstData As Recordset) 
 
Dim qdfData As DAO.QueryDef 
Dim strNewFilter As String 
Dim strRightSQL As String 
 
On Error GoTo ErrorHandler 
 
   Set qdfData = rstData.CopyQueryDef 
 
   ' Try "LastName LIKE 'D*'". 
   strNewFilter = InputBox("Enter new criteria") 
 
   strRightSQL = Right(qdfData.SQL, 1) 
 
   ' Strip characters from the end of the query, 
   ' as needed. 
   Do While strRightSQL = " " Or strRightSQL = ";" Or _ 
                          strRightSQL = vbCR Or strRightSQL = vbLF 
      qdfData.SQL = Left(qdfData.SQL, Len(qdfData.SQL) - 1) 
      strRightSQL = Right(qdfData.SQL, 1) 
   Loop 
 
   qdfData.SQL = qdfData.SQL &; " AND " &; strNewFilter 
   rstData.Requery qdfData         'Requery the Recordset. 
   rstData.MoveLast               'Populate the Recordset. 
 
   ' "Lastname LIKE 'D*'" should return 2 records. 
   MsgBox "Number of records found:  " &; rstData.RecordCount &; "." 
 
   qdfData.Close 
   Set qdfData = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub
```


 **Note**  To use the  **Requery** method, the **[Restartable](http://msdn.microsoft.com/library/00DEF49D-EA7E-6CD5-2F4A-914A1DDCDD51%28Office.15%29.aspx)** property of the **Recordset** object must be set to **True**. The **Restartable** property is always set to **True** when the **Recordset** is created from a query other than a crosstab query against tables in an Access database. You cannot restart SQL pass-through queries. You may or may not be able to restart queries against linked tables in another database format. To determine whether a **Recordset** object can rerun its query, check the **Restartable** property.


