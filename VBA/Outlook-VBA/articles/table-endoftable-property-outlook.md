---
title: Table.EndOfTable Property (Outlook)
keywords: vbaol11.chm2237
f1_keywords:
- vbaol11.chm2237
ms.prod: outlook
api_name:
- Outlook.Table.EndOfTable
ms.assetid: 8c185230-65ce-1b66-7b63-8de3533dea86
ms.date: 06/08/2017
---


# Table.EndOfTable Property (Outlook)

Returns a  **Boolean** that indicates whether the current row is positioned after the last row in the **[Table](table-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **EndOfTable**

 _expression_ A variable that represents a **Table** object.


## Remarks

When you open a  **Table** using **[Folder.GetTable](folder-gettable-method-outlook.md)** , **[Search.GetTable](search-gettable-method-outlook.md)** , or **[Table.Restrict](table-restrict-method-outlook.md)** , the value of **EndOfTable** depends on the number of rows in the **Table** . If there are no rows, **EndOfTable** returns **True**. If there are rows,  **EndOfTable** returns **False** until the cursor moves beyond the last row in the **Table**. 

The  **Table** has two virtual **Null** rows, one before the first row and one after the last row in the **Table** . Each call to **[Table.GetNextRow](table-getnextrow-method-outlook.md)** moves the current row to the next row in the **Table** and returns a **[Row](row-object-outlook.md)** object that represents the current row.

The  **EndOfTable** property returns **True** if the current row is after the last row and **False** if the current row is on or before the last row.


## Example

The following code sample applies a Jet filter on items in the Inbox to obtain a  **Table** of items that have a **LastModificationTime** greater than November 1, 2005. It then uses the **EndOfTable** property to enumerate the items in the **Table** and print the values of the default properties for each item in the **Table** .


```vb
Sub DemoTable() 
 'Declarations 
 Dim Filter As String 
 Dim oRow As Outlook.Row 
 Dim oTable As Outlook.Table 
 Dim oFolder As Outlook.Folder 
 
 'Get a Folder object for the Inbox 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 'Define Filter to obtain items last modified after November 1, 2005 
 Filter = "[LastModificationTime] > '11/1/2005'" 
 'Restrict with Filter 
 Set oTable = oFolder.GetTable(Filter) 
 
 'Enumerate the table using test for EndOfTable 
 Do Until (oTable.EndOfTable) 
 Set oRow = oTable.GetNextRow() 
 Debug.Print (oRow("EntryID")) 
 Debug.Print (oRow("Subject")) 
 Debug.Print (oRow("CreationTime")) 
 Debug.Print (oRow("LastModificationTime")) 
 Debug.Print (oRow("MessageClass")) 
 Loop 
End Sub
```


## See also


#### Concepts


[Table Object](table-object-outlook.md)

