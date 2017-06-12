---
title: Create a DAO Recordset From a Query
ms.prod: access
ms.assetid: d84870d4-58e4-9d48-9951-72d928929002
ms.date: 06/08/2017
---


# Create a DAO Recordset From a Query

You can create a  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object based on a stored select query. In the following code example, Current Product List is an existing select query stored in the current database.


```vb
Dim dbsNorthwind As DAO.Database 
Dim rstProducts As DAO.Recordset 
 
Set dbsNorthwind = CurrentDb 
Set rstProducts = dbsNorthwind.OpenRecordset("Current Product List") 

```


If a stored select query does not already exist, the  **[OpenRecordset](http://msdn.microsoft.com/library/7D5CA4D5-5A0B-C0C8-D8E8-2C4E6C5F361F%28Office.15%29.aspx)** method also accepts an SQL string instead of the name of a query. The previous example can be rewritten as follows.




```vb
Dim dbsNorthwind As DAO.Database 
Dim rstProducts As DAO.Recordset 
Dim strSQL As String 
 
Set dbsNorthwind = CurrentDb 
strSQL = "SELECT * FROM Products WHERE Discontinued = No " &; _ 
         "ORDER BY ProductName" 
Set rstProducts = dbsNorthwind.OpenRecordset(strSQL) 

```

The disadvantage of this approach is that the query string must be compiled each time it runs, whereas the stored query is compiled the first time it is saved, which usually results in slightly better performance.

