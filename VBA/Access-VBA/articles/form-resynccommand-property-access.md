---
title: Form.ResyncCommand Property (Access)
keywords: vbaac10.chm13486
f1_keywords:
- vbaac10.chm13486
ms.prod: access
api_name:
- Access.Form.ResyncCommand
ms.assetid: 0df53ea9-5771-0ccd-07ef-f33ad1082a61
ms.date: 06/08/2017
---


# Form.ResyncCommand Property (Access)

You can use the  **ResyncCommand** property to specify or determine the SQL statement or stored procedure that will be used in an updateable snapshot of a table. Read/write **String**.


## Syntax

 _expression_. **ResyncCommand**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **ResyncCommand** property is a string expression representing a SQL statement or stored procedure that is parameterized by the key columns from the Unique Table in the output cursor, using ? as parameter markers.

The parameters must match in number and ordering to the set of key columns for the table identified by the  **[UniqueTable](form-uniquetable-property-access.md)** property. The purpose of the **ResyncCommand** property is to pull in the "fixed up" values of a row in a recordset after an update has been made, including an update to a join column.

For data access pages and for forms based on views or non-parameterized SQL queries containing a join, if the  **ResyncCommand** property is null, Microsoft Access determines an appropriate query to use for the resync operation. For data access pages and forms based on stored procedures or parameterized SQL statements, Access cannot determine an appropriate resync query at run time, so the user must supply the **ResyncCommand** string in order to get the correct row fix up behavior. If the **ResyncCommand** property is empty and Access cannot determine an appropriate query to use, the default ADO resync operation (to display the current values) occurs after an update or insert.


## See also


#### Concepts


[Form Object](form-object-access.md)

