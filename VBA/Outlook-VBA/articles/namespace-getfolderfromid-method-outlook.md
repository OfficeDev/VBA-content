---
title: NameSpace.GetFolderFromID Method (Outlook)
keywords: vbaol11.chm762
f1_keywords:
- vbaol11.chm762
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetFolderFromID
ms.assetid: 0fb2d3b5-2967-1943-922a-7ec03e514e62
ms.date: 06/08/2017
---


# NameSpace.GetFolderFromID Method (Outlook)

Returns a  **[Folder](folder-object-outlook.md)** object identified by the specified entry ID (if valid).


## Syntax

 _expression_ . **GetFolderFromID**( **_EntryIDFolder_** , **_EntryIDStore_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EntryIDFolder_|Required| **String**|The  **[EntryID](folder-entryid-property-outlook.md)** of the folder.|
| _EntryIDStore_|Optional| **Variant**|The  **[StoreID](folder-storeid-property-outlook.md)** for the folder.|

### Return Value

A ** Folder** object that represents the specified folder.


## Remarks

This method is used for ease of transition between MAPI and OLE/Messaging applications and Microsoft Outlook.


## Example

This Visual Basic for Applications (VBA) example obtains the  **EntryID** and **StoreID** for the default **Tasks** folder and then calls the **GetFolderFromID** method using these values to obtain the same folder. The folder is then displayed.


```vb
Sub GetWithID() 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myEntryID As String 
 
 Dim myStoreID As String 
 
 Dim myNewFolder As Outlook.Folder 
 
 
 
 Set myFolder = Application.Session.GetDefaultFolder(olFolderTasks) 
 
 myEntryID = myFolder.EntryID 
 
 myStoreID = myFolder.StoreID 
 
 Set myNewFolder = Application.Session.GetFolderFromID(myEntryID, myStoreID) 
 
 myNewFolder.Display 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

