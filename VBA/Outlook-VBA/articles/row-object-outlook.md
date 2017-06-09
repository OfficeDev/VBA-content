---
title: Row Object (Outlook)
keywords: vbaol11.chm3167
f1_keywords:
- vbaol11.chm3167
ms.prod: outlook
api_name:
- Outlook.Row
ms.assetid: 06db3fa4-1649-48bf-3b86-ffdf99a47305
ms.date: 06/08/2017
---


# Row Object (Outlook)

Represents a row of data in the  **[Table](table-object-outlook.md)** object.


## Remarks

A  **Table** is composed of rows and columns. It represents a read-only dynamic rowset of data in a **[Folder](folder-object-outlook.md)** or **[Search](search-object-outlook.md)** object. You can regard each row of a **Table** as an item in the folder, and each column as a property of the item. By default, the **Table** contains only a subset of properties for items in the folder. This makes the **Table** an in-memory lightweight rowset that supports fast enumeration and filtering of items in the folder.

 If the **Table** object is obtained from **[Folder.GetTable](folder-gettable-method-outlook.md)**, you can further specify a filter (in **[Table.Restrict](table-restrict-method-outlook.md)** ) to obtain a more restricted set of rows in the **Table**.

 You can use the Table methods: **[FindRow](table-findrow-method-outlook.md)**, **[FindNextRow](table-findnextrow-method-outlook.md)**, **[GetNextRow](table-getnextrow-method-outlook.md)**, and **[MoveToStart](table-movetostart-method-outlook.md)** to obtain a specific row in a **Table**.

 Use **[Row.GetValues](row-getvalues-method-outlook.md)** to obtain an array of values that correspond to column values at that row in the **Table**.

 Use the helper functions **[Row.BinaryToString](row-binarytostring-method-outlook.md)**, **[Row.LocalTimeToUTC](row-localtimetoutc-method-outlook.md)**, and **[Row.UTCToLocalTime](row-utctolocaltime-method-outlook.md)** to facilitate type conversion of column values at a specific row. For more information on property value representation in a **Table**, see[Factors Affecting Property Value Representation in the Table and View Classes](http://msdn.microsoft.com/library/13cf9945-a9e0-bb32-a2cb-74366a365ae1%28Office.15%29.aspx).

 Although additions and deletions of the underlying folder are reflected by the rows in the **Table**, the **Table** does not support any events for adding, changing, and removing of rows. If you require a writeable object from the **Table** row, obtain the Entry ID for that row from the default EntryID column in the **Table** and then use the **[GetItemFromID](namespace-getitemfromid-method-outlook.md)** method of the **[NameSpace](namespace-object-outlook.md)** object to obtain a full item, such as a **[MailItem](mailitem-object-outlook.md)** or **[ContactItem](contactitem-object-outlook.md)**, that supports read-write operations. For more information on default columns in a **Table**, see[Default Properties Displayed in a Table Object](http://msdn.microsoft.com/library/649c64f3-2d1e-23f1-bf13-3368da79e62b%28Office.15%29.aspx).


## Example

The following code sample illustrates how to obtain a  **Table** object based on the **LastModificationTime** of items in the Inbox. It also shows how to customize columns in the **Table**, and how to enumerate and print the values of the corresponding properties of these items.


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
|[BinaryToString](row-binarytostring-method-outlook.md)|
|[GetValues](row-getvalues-method-outlook.md)|
|[Item](row-item-method-outlook.md)|
|[LocalTimeToUTC](row-localtimetoutc-method-outlook.md)|
|[UTCToLocalTime](row-utctolocaltime-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](row-application-property-outlook.md)|
|[Class](row-class-property-outlook.md)|
|[Parent](row-parent-property-outlook.md)|
|[Session](row-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
