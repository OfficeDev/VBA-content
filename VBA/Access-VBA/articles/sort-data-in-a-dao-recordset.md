---
title: Sort Data in a DAO Recordset
ms.prod: access
ms.assetid: 900b0b00-34f5-dba6-5386-34360cee95a0
ms.date: 06/08/2017
---


# Sort Data in a DAO Recordset

Unless you open a table-type  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object and set its **Index** property, you cannot be sure that records will appear in any specific order. However, you usually want to retrieve records in a specific order. For example, you may want to view invoices arranged by increasing invoice number, or retrieve employee records in alphabetical order by their last names. To see records in a specific order, sort them.

To sort data in a  **Recordset** object that is not a table, use an SQL ORDER BY clause in the query that constructs the **Recordset**. You can specify an SQL string when you create a **[QueryDef](http://msdn.microsoft.com/library/0B3D901C-345D-42A2-F5F1-FB09CC562E27%28Office.15%29.aspx)** object, when you create a stored query in a database, or when you use the **[OpenRecordset](http://msdn.microsoft.com/library/B4908C36-C156-E269-E2AD-B1FA20EC4884%28Office.15%29.aspx)** method.

You can also filter data, which means you restrict the result set returned by a query to records that meet some criteria. With any type of  **OpenRecordset** object, use an SQL WHERE clause in the original query to filter data.

The following code example opens a dynaset-type  **Recordset** object, and uses an SQL statement to retrieve, filter, and sort records.



```vb
Dim dbsNorthwind As DAO.Database 
Dim rstManagers As DAO.Recordset 
 
Set dbsNorthwind = CurrentDb 
Set rstManagers = dbsNorthwind.OpenRecordset("SELECT FirstName, " &; _ 
                  "LastName FROM Employees WHERE Title = " &; _ 
                  "'Sales Manager' ORDER BY LastName") 

```

One limitation of running an SQL query in an  **OpenRecordset** method is that it has to be recompiled every time you run it. If this query is used frequently, you can improve performance by first creating a stored query using the same SQL statement, and then opening a **Recordset** object against the query, as shown in the following code example.



```vb
Dim dbsNorthwind As DAO.Database 
Dim rstSalesReps As DAO.Recordset 
Dim qdfSalesReps As DAO.QueryDef 
 
Set dbsNorthwind = CurrentDb 
 
Set qdfSalesReps = dbsNorthwind.CreateQueryDef("SalesRepQuery") 
qdfSalesReps.SQL = "SELECT * FROM Employees WHERE Title = " &; _ 
                   "'Sales Representative'" 
 
Set rstSalesReps = dbsNorthwind.OpenRecordset("SalesRepQuery") 

```


