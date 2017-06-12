---
title: Application.CodeDb Method (Access)
keywords: vbaac10.chm12547
f1_keywords:
- vbaac10.chm12547
ms.prod: access
api_name:
- Access.Application.CodeDb
ms.assetid: 7f0cff23-1265-231f-9ab5-fa83c19d39cf
ms.date: 06/08/2017
---


# Application.CodeDb Method (Access)

You can use the  **CodeDb** method in a code module to determine the name of the **Database** object that refers to the database in which code is currently running. Use the **CodeDb** method to access Data Access Objects (DAO) that are part of a library database.


## Syntax

 _expression_. **CodeDb**

 _expression_ A variable that represents an **Application** object.


### Return Value

Database


## Remarks

For example, you can use the  **CodeDb** method in a module in a library database to create a **Database** object referring to the library database. You can then open a recordset based on a table in the library database.

 **Set** _database_=  **CodeDb**

The  **CodeDb** method returns a **Database** object for which the **Name** property is the full path and name of the database from which it is called. This method can be useful when you need to manipulate the Data Access Objects in your library database.

When you call a method in a library database, the database from which you have called the method remains the current database, even while code is running in a module in the library database. In order to refer to the Data Access Objects in the library database, you need to know the name of the  **Database** object that represents the library database.

For example, suppose you have a table in a library database that lists error messages. To manipulate data in the table from code, you could use the  **CodeDb** method to determine the name of the **Database** object that refers to the library database that contains the table.

If the  **CodeDb** method is run from the current database, it returns the name of the current database, which is the same value returned by the **[CurrentDb](application-currentdb-method-access.md)** method.


## Example

The following example uses the  **CodeDb** method to return a **Database** object that refers to a library database. The library database contains both a table named Errors and the code that is currently running. After the **CodeDb** method determines this information, the GetErrorString function opens a table-type recordset based on the Errors table. It then extracts an error message from a field named ErrorData based on the **Integer** value passed to the function.


```vb
Function GetErrorString(ByVal intError As Integer) As String 
 Dim dbs As Database, rst As RecordSet 
 
 ' Variable refers to database where code is running. 
 Set dbs = CodeDb 
 ' Create table-type Recordset object. 
 Set rst = dbs.OpenRecordSet("Errors", dbOpenTable) 
 ' Set index to primary key (ErrorID field). 
 rst.Index = "PrimaryKey" 
 ' Find error number passed to GetErrorString function. 
 rst.Seek "=", intError 
 ' Return associated error message. 
 GetErrorString = rst.Fields!ErrorData.Value 
 rst.Close 
End Function
```


## See also


#### Concepts


[Application Object](application-object-access.md)

