---
title: Table.Restrict Method (Outlook)
keywords: vbaol11.chm2234
f1_keywords:
- vbaol11.chm2234
ms.prod: outlook
api_name:
- Outlook.Table.Restrict
ms.assetid: ecdd30f6-e12c-8025-3ded-592d2fad2bb8
ms.date: 06/08/2017
---


# Table.Restrict Method (Outlook)

Applies a filter to the rows in the  **[Table](table-object-outlook.md)** and obtains a new **Table** object.


## Syntax

 _expression_ . **Restrict**( **_Filter_** )

 _expression_ A variable that represents a **Table** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filter_|Required| **String**|Specifies the criteria for rows in the  **Table** object.|

### Return Value

A  **Table** object that is returned by applying _Filter_ to the rows in the parent **Table** object.


## Remarks

You can only use  **Table.Restrict** to apply another filter to that **Table** if the parent object of the **Table** is a **[Folder](folder-object-outlook.md)** object. If the parent object is a **[Search](search-object-outlook.md)** object, **Restrict** will return an error.

Since the filter is applied to the rows in the  **Table** object, this is equivalent to applying a filter that is a logical **AND** of _Filter_ and all preceding filters applied to the same **Table** object.

 _Filter_ is a query on specified properties of items that are represented as rows in the parent **Table** . The query uses either the Microsoft Jet syntax or the DAV Searching and Locating (DASL) syntax. For example, the following Jet filter and DASL filter specify the same criteria for items with **LastModificationTime** earlier than 3:30pm of June 12, 2005:




```
criteria = "[LastModificationTime] < '" &; Format$("6/12/2005 3:30PM","General Date") &; "'"criteria = "@SQL=" &; Chr(34) &; "DAV:getlastmodified" &; Chr(34) 
 &; " < '" &; Format$("6/12/2005 3:30PM","General Date") &; "'" 

```

For more information on specifying filters for the  **Table** object, see[Filtering Items](http://msdn.microsoft.com/library/4038e042-1b07-5d18-18b0-c2b58c9c42da%28Office.15%29.aspx).

If  _Filter_ contains custom properties, those properties must exist in the parent folder of the **Table** object in order for the restriction to work correctly. Certain properties are not supported in a **Table** filter, including binary properties, computed properties, and HTML or RTF body content. For more information, see[Unsupported Properties in a Table Object or Table Filter](http://msdn.microsoft.com/library/0e37f03f-7677-ca29-d0b2-8b45c026e5f1%28Office.15%29.aspx).

If  _Filter_ is a blank string, **Restrict** returns a **Table** object that is identical to the parent **Table** object.


## Example

The following code sample applies a Jet filter on items in the Inbox to obtain items with a  **LastModificationTime** greater than November 1, 2005. It then prints the values of the default properties for these items in the Inbox: **EntryID** , **Subject** , **CreationTime** , **LastModificationTime** , and **MessageClass** .


 **Note**  Since heterogeneous items can exist in the same folder in Outlook, the items returned from applying the filter to the Inbox may be of different types. In general, before accessing any properties that are not among the default properties for items in the Inbox, you should check the  **MessageClass** of the item.


```vb
Sub RestrictTable() 
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

