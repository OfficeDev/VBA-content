---
title: Folder.GetTable Method (Outlook)
keywords: vbaol11.chm2018
f1_keywords:
- vbaol11.chm2018
ms.prod: outlook
api_name:
- Outlook.Folder.GetTable
ms.assetid: 08d184cb-0c41-01b1-abc5-305476380f8b
ms.date: 06/08/2017
---


# Folder.GetTable Method (Outlook)

Obtains a  **[Table](table-object-outlook.md)** object that contains items filtered by _Filter_ .


## Syntax

 _expression_ . **GetTable**( **_Filter_** , **_TableContents_** )

 _expression_ A variable that represents a **[Folder](folder-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filter_|Optional| **String**|A filter in Microsft Jet or DAV Searching and Locating (DASL) syntax that specifies the criteria for items in the parent  **Folder** .|
| _TableContents_|Optional| **[OlTableContents](oltablecontents-enumeration-outlook.md)**|Specifies the type of items in the folder that  **GetTable** returns. The default is **olUserItems**.|

### Return Value

A  **Table** that contains items in the parent **[Folder](folder-object-outlook.md)** that meet the criteria in _Filter_ . By default, _TableContents_ is **olUserItems** and the returned **Table** contains only the filtered items that are not hidden.


## Remarks

If  _Filter_ is a blank string or the _Filter_ parameter is omitted, **GetTable** returns a **Table** with rows representing all the items in the **Folder** . If _Filter_ is a blank string or the _Filter_ parameter is omitted and _TableContents_ is **olHiddenItems** , **GetTable** returns a **Table** with rows representing all the hidden items in the **Folder** .

For more information on filters, see [Filtering Items](http://msdn.microsoft.com/library/4038e042-1b07-5d18-18b0-c2b58c9c42da%28Office.15%29.aspx) and[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).

 **GetTable** returns a **Table** with the default column set for the folder type of the parent **Folder** . To modify the default column set, use the **[Add](columns-add-method-outlook.md)** , **[Remove](columns-remove-method-outlook.md)** , or **[RemoveAll](columns-removeall-method-outlook.md)** methods of the **[Columns](columns-object-outlook.md)** collection object. When _TableContents_ is **olHiddenItems** , the default column set is always the default column set for a mail folder even though the parent **Folder** might be, for example, a Contacts folder. For more information on default column sets, see[Default Properties Displayed in a Table Object](http://msdn.microsoft.com/library/649c64f3-2d1e-23f1-bf13-3368da79e62b%28Office.15%29.aspx).

You can use  **[Table.Restrict](table-restrict-method-outlook.md)** to apply subsequent filters to a **Table** that is based on the **Folder** object.


## Example

The following code sample illustrates how to use  **Folder.GetTable** to obtain a **Table** object based on the **LastModificationTime** of items in the Inbox. It then enumerates and prints the values of a couple of default properties of these items.


```vb
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
  
    'Enumerate the table using test for EndOfTable  
    Do Until (oTable.EndOfTable)  
        Set oRow = oTable.GetNextRow()  
        Debug.Print (oRow("Subject"))  
        Debug.Print (oRow("LastModificationTime"))  
    Loop  
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

