---
title: Folders.FolderAdd Event (Outlook)
keywords: vbaol11.chm308
f1_keywords:
- vbaol11.chm308
ms.prod: outlook
api_name:
- Outlook.Folders.FolderAdd
ms.assetid: d72beffe-5a6b-41f1-0a0e-2f8548cbdc84
ms.date: 06/08/2017
---


# Folders.FolderAdd Event (Outlook)

Occurs when a folder is added to the specified  **[Folders](folders-object-outlook.md)** collection.


## Syntax

 _expression_ . **FolderAdd**( **_Folder_** )

 _expression_ A variable that represents a **Folders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Folder_|Required| **[Folder](folder-object-outlook.md)**|The  **Folder** that is added.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays a new folder created in the user?s  **Inbox** folder.


```vb
Public WithEvents myOlFolders As Outlook.Folders 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlFolders = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders 
 
End Sub 
 
 
 
Private Sub myOlFolders_FolderAdd(ByVal Folder As Outlook.Folder) 
 
 Folder.Display 
 
End Sub
```


## See also


#### Concepts


[Folders Object](folders-object-outlook.md)

