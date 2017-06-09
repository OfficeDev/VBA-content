---
title: Folders.FolderChange Event (Outlook)
keywords: vbaol11.chm309
f1_keywords:
- vbaol11.chm309
ms.prod: outlook
api_name:
- Outlook.Folders.FolderChange
ms.assetid: cd379b87-6fb7-bfa4-544a-0c406a170832
ms.date: 06/08/2017
---


# Folders.FolderChange Event (Outlook)

Occurs when a folder in the specified  **[Folders](folders-object-outlook.md)** collection is changed.


## Syntax

 _expression_ . **FolderChange**( **_Folder_** )

 _expression_ A variable that represents a **Folders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Folder_|Required| **[Folder](folder-object-outlook.md)**|The folder that has been changed.|

## Remarks

The  **FolderChange** event fires when a folder in a **Folders** collection object is changed, either through user action or program code. The change can be a user or program code renaming the folder, or adding, changing, or removing an item in the folder. This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example prompts the user to remove a folder from the  **Deleted Items** folder if the folder is empty. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents myFolders As Outlook.Folders 
 
 
 
Sub Initialize_handler() 
 
 Set myNS = Application.GetNamespace("MAPI") 
 
 Set myFolders = myNS.GetDefaultFolder(olFolderDeletedItems).Folders 
 
End Sub 
 
 
 
Private Sub myFolders_FolderChange(ByVal Folder As Outlook.Folder) 
 
 If Folder.Items.Count = 0 Then 
 
 MyPrompt = Folder.Name &; " is empty. Do you want to delete it?" 
 
 If MsgBox(MyPrompt, vbYesNo + vbQuestion) = vbYes Then 
 
 Folder.Delete 
 
 End If 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Folders Object](folders-object-outlook.md)

