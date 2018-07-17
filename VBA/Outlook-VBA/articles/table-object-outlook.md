---
title: Table Object (Outlook)
keywords: vbaol11.chm3166
f1_keywords:
- vbaol11.chm3166
ms.prod: outlook
api_name:
- Outlook.Table
ms.assetid: 0affaafd-93fe-227a-acee-e09a86cadc20
ms.date: 06/08/2017
---


# Table Object (Outlook)

Represents a set of item data from a  **[Folder](folder-object-outlook.md)** or **[Search](search-object-outlook.md)** object, with items as rows of the table and properties as columns of the table.


## Remarks

The  **Table** represents a read-only dynamic rowset of data in a **Folder** or **Search** object. You can use **[Folder.GetTable](http://msdn.microsoft.com/library/08d184cb-0c41-01b1-abc5-305476380f8b%28Office.15%29.aspx)** or **[Search.GetTable](http://msdn.microsoft.com/library/3aba6b77-73a3-9620-9c18-b2e03c7b63bc%28Office.15%29.aspx)** to obtain a **Table** object that represents a set of items in a folder or search folder. If the **Table** object is obtained from **Folder.GetTable**, you can further specify a filter (in **[Table.Restrict](http://msdn.microsoft.com/library/ecdd30f6-e12c-8025-3ded-592d2fad2bb8%28Office.15%29.aspx)** ) to obtain a subset of the items in the folder. If you do not specify any filter, you will obtain all the items in the folder.

By default, each item in the returned  **Table** contains only a default subset of its properties. You can regard each row of a **Table** as an item in the folder, each column as a property of the item, and the **Table** as an in-memory lightweight rowset that allows fast enumeration and filtering of items in the folder. Although additions and deletions of the underlying folder are reflected by the rows in the **Table**, the **Table** does not support any events for adding, changing, and removing of rows. If you require a writeable object from the **Table** row, obtain the Entry ID for that row from the default EntryID column in the **Table** and then use the **[GetItemFromID](http://msdn.microsoft.com/library/f2abff80-4c04-998b-654b-28600424a16f%28Office.15%29.aspx)** method of the **[NameSpace](namespace-object-outlook.md)** object to obtain a full item, such as a **[MailItem](http://msdn.microsoft.com/library/14197346-05d2-0250-fa4c-4a6b07daf25f%28Office.15%29.aspx)** or **[ContactItem](contactitem-object-outlook.md)**, that supports read-write operations. For more information on default columns in a **Table**, see[Default Properties Displayed in a Table Object](http://msdn.microsoft.com/library/649c64f3-2d1e-23f1-bf13-3368da79e62b%28Office.15%29.aspx).

 For more information on the **Table** object, see[Enumerating, Searching, and Filtering Items in a Folder](http://msdn.microsoft.com/library/d786d292-7a0e-0e1a-e132-affbfde37744%28Office.15%29.aspx).


## Example

The following code sample illustrates how the  **Table** object can return a filtered set of items based on their **LastModificationTime** property. It also shows how to list the default properties as well as specific properties of the items.


```
Sub DemoTable() 
 
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


## Methods



|**Name**|
|:-----|
|[FindNextRow](http://msdn.microsoft.com/library/e09019ca-e4bb-2597-7b9e-a56c1b5fce6c%28Office.15%29.aspx)|
|[FindRow](http://msdn.microsoft.com/library/5722cf58-d026-007a-558f-90b73bad920d%28Office.15%29.aspx)|
|[GetArray](http://msdn.microsoft.com/library/2594bb2e-290f-8e88-52d1-cd2b2191bbe3%28Office.15%29.aspx)|
|[GetNextRow](http://msdn.microsoft.com/library/e01ddaa0-a869-2f52-5e46-84d4d4090e61%28Office.15%29.aspx)|
|[GetRowCount](http://msdn.microsoft.com/library/06014c43-700a-8502-bad7-b3f93a22e870%28Office.15%29.aspx)|
|[MoveToStart](http://msdn.microsoft.com/library/af499471-dd21-9374-7399-3ce977368015%28Office.15%29.aspx)|
|[Restrict](http://msdn.microsoft.com/library/ecdd30f6-e12c-8025-3ded-592d2fad2bb8%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/4e4867c2-27b8-f920-59ce-b60116d22054%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/10e7611e-e3b3-a07c-da85-f8c270a37212%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/bea314b0-9db9-ac67-a897-49e619da1066%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/57005ab1-ad49-296d-5b34-24dfd8f0987f%28Office.15%29.aspx)|
|[EndOfTable](http://msdn.microsoft.com/library/8c185230-65ce-1b66-7b63-8de3533dea86%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1c6a54ac-ba4d-72a2-0871-a3522582dbde%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/8a17876d-6637-f30b-6c0f-32cfc8b77d51%28Office.15%29.aspx)|

## See also


#### Other resources


[Table Object Members](http://msdn.microsoft.com/library/bd9db35d-0738-22cf-a936-425d5a0ead87%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
