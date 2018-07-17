---
title: CodeData Object (Access)
keywords: vbaac10.chm12742
f1_keywords:
- vbaac10.chm12742
ms.prod: access
api_name:
- Access.CodeData
ms.assetid: fc207136-4d18-2c7d-ffe6-0e1ad7c2fc32
ms.date: 06/08/2017
---


# CodeData Object (Access)

The  **CodeData** object refers to objects stored within the code database by the source (server) application.


## Remarks

The  **CodeData** object has several collections that contain specific object types within the code database. The following table lists the name of each collection defined by the database and the types of objects it contains.



|**Collections**|**Object type**|**Available in Access database **|**Available in Access Project (.adp)**|
|:-----|:-----|:-----|:-----|
|**[AllTables](alltables-object-access.md)**|All tables|Yes|Yes|
|**[AllFunctions](allfunctions-object-access.md)**|All functions|No|Yes|
|**[AllQueries](allqueries-object-access.md)**|All queries |Yes|Yes|
|**[AllViews](allviews-object-access.md)**|All views |No|Yes|
|**[AllStoredProcedures](allstoredprocedures-object-access.md)**|All stored procedures |No|Yes|
|**[AllDatabaseDiagrams](alldatabasediagrams-object-access.md)**|All database diagrams |No|Yes|

 **Note**  The collections in the preceding table contain all of the respective objects in the database regardless if they are opened or closed.

For example, an  **AccessObject** representing a table is a member of the **AllTables** collection, which is a collection of **AccessObject** objects within the current database. Within the **AllTables** collection, individual tables are indexed beginning with zero. You can refer to an individual **AccessObject** object in the **AllTables** collection either by referring to the table by name, or by referring to its index within the collection. If you want to refer to a specific item in the **AllTables** collection, it's better to refer to it by name because the item's index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**AllTables** ! _tablename_|AllTables!OrderTable|
|**AllTables** ![ _table name_]|AllTables![Order Table]|
|**AllTables** (" _tablename_")|AllTables("OrderTable")|
|**AllTables** ( _index_)|AllTables(0)|

## Properties



|**Name**|
|:-----|
|[AllDatabaseDiagrams](codedata-alldatabasediagrams-property-access.md)|
|[AllFunctions](codedata-allfunctions-property-access.md)|
|[AllQueries](codedata-allqueries-property-access.md)|
|[AllStoredProcedures](codedata-allstoredprocedures-property-access.md)|
|[AllTables](codedata-alltables-property-access.md)|
|[AllViews](codedata-allviews-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
