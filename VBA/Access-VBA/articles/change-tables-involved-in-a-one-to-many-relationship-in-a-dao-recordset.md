---
title: Change Tables Involved in a One-to-Many Relationship in a DAO Recordset
ms.prod: access
ms.assetid: d859066f-dfb5-e3b5-0cc8-fed38c3adb26
ms.date: 06/08/2017
---


# Change Tables Involved in a One-to-Many Relationship in a DAO Recordset

Dynaset-type  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** objects can be based on a multiple-table query containing tables with a one-to-many relationship. For example, suppose you want to create a multiple-table query that combines fields from the Orders and Order Details tables. Generally speaking, you cannot change values in the Orders table because it is on the "one" side of the relationship. Depending on your application, however, you may want to be able to make changes to the Orders table.

To make it possible to freely change the values on the "one" side of a one-to-many relationship, use the  **dbInconsistent** constant of the **[OpenRecordset](http://msdn.microsoft.com/library/7D5CA4D5-5A0B-C0C8-D8E8-2C4E6C5F361F%28Office.15%29.aspx)** method to create an inconsistent dynaset. For example.



```vb
Set rstTotalSales = dbs.OpenRecordset("Sales Totals" ,,dbInconsistent)
```

When you update an inconsistent dynaset, you can easily destroy the referential integrity of the data in the dynaset. You must take care to understand how the data is related across the one-to-many relationship and to update the values on both sides in a way that preserves data integrity.
The  **dbInconsistent** constant is available only for dynaset-type **Recordset** objects. It is ignored for table, snapshot, and forward-only-type **Recordset** objects, but no compile or run-time error is returned if the **dbInconsistent** constant is used with those types of **Recordset** objects.
Even with an inconsistent  **Recordset**, some fields may not be updatable. For example, you cannot change the value of an AutoNumber field, and a **Recordset** based on certain linked tables may not be updatable.

