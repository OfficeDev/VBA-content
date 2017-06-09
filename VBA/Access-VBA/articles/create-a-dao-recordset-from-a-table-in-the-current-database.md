---
title: Create a DAO Recordset From a Table In the Current Database
ms.prod: access
ms.assetid: b0507965-e6af-cda4-9d50-fbeb98b4ab89
ms.date: 06/08/2017
---


# Create a DAO Recordset From a Table In the Current Database

The following code example uses the  **[OpenRecordset](http://msdn.microsoft.com/library/7D5CA4D5-5A0B-C0C8-D8E8-2C4E6C5F361F%28Office.15%29.aspx)** method to create a table-type **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object for a table in the current database.


```vb
Dim dbsNorthwind As DAO.Database 
Dim rstCustomers As DAO.Recordset 
 
Set dbsNorthwind = CurrentDb 
Set rstCustomers = dbsNorthwind.OpenRecordset("Customers") 

```


