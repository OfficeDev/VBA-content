---
title: Columns.RemoveAll Method (Outlook)
keywords: vbaol11.chm2743
f1_keywords:
- vbaol11.chm2743
ms.prod: outlook
api_name:
- Outlook.Columns.RemoveAll
ms.assetid: e9923548-9c75-e5dd-0643-3c42cd112352
ms.date: 06/08/2017
---


# Columns.RemoveAll Method (Outlook)

Removes all the columns from the  **[Columns](columns-object-outlook.md)** collection and resets the **[Table](table-object-outlook.md)** .


## Syntax

 _expression_ . **RemoveAll**

 _expression_ A variable that represents a **Columns** object.


## Remarks

 **RemoveAll** resets the **Table** by moving the current row to just before the first row of the **Table** . After a call to **RemoveAll** , **[Columns.Count](columns-count-property-outlook.md)** becomes zero (0).


## Example

The following code sample illustrates how to obtain a  **Table** object based on the **LastModificationTime** of items in the Inbox. It also shows how to remove the default columns of the **Table** , add specific columns, and print the values of the corresponding properties of these items.


```vb
Sub RemoveAllAndAddColumns() 
 
 'Declarations 
 
 Dim Filter As String 
 
 Dim oRow As Outlook.Row 
 
 Dim oTable As Outlook.Table 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 'Get a Folder object for the Inbox 
 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 
 
 'Define Filter to obtain items last modified after May 1, 2005 
 
 Filter = "[LastModificationTime] > '5/1/2005'" 
 
 'Restrict with Filter 
 
 Set oTable = oFolder.GetTable(Filter) 
 
 
 
 'Remove all columns in the default column set 
 
 oTable.Columns.RemoveAll 
 
 'Specify desired properties 
 
 With oTable.Columns 
 
 .Add ("Subject") 
 
 .Add ("LastModificationTime") 
 
 'PR_ATTR_HIDDEN referenced by the MAPI proptag namespace 
 
 .Add ("http://schemas.microsoft.com/mapi/proptag/0x10F4000B") 
 
 End With 
 
 
 
 'Enumerate the table using test for EndOfTable 
 
 Do Until (oTable.EndOfTable) 
 
 Set oRow = oTable.GetNextRow() 
 
 Debug.Print (oRow("Subject")) 
 
 Debug.Print (oRow("LastModificationTime")) 
 
 Debug.Print (oRow("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) 
 
 Loop 
 
End Sub
```


## See also


#### Concepts


[Columns Object](columns-object-outlook.md)

