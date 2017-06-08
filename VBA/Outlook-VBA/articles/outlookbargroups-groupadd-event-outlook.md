---
title: OutlookBarGroups.GroupAdd Event (Outlook)
keywords: vbaol11.chm355
f1_keywords:
- vbaol11.chm355
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroups.GroupAdd
ms.assetid: 5fae2579-b4db-d645-27d4-dce867e64242
ms.date: 06/08/2017
---


# OutlookBarGroups.GroupAdd Event (Outlook)

Occurs when a new group has been added to the  **Shortcuts** pane.


## Syntax

 _expression_ . **GroupAdd**( **_NewGroup_** )

 _expression_ A variable that represents an **OutlookBarGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewGroup_|Required| **[OutlookBarGroup](outlookbargroup-object-outlook.md)**|The  **OutlookBarGroup** that was added.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example adds a shortcut to the  **Calendar** whenever a group is created. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents myOlGroups As Outlook.OutlookBarGroups 
Dim myOlBar As Outlook.OutlookBarPane 
 
Sub Initialize_handler() 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 Set myOlGroups = myOlBar.Contents.Groups 
End Sub 
 
Private Sub myOlGroups_GroupAdd(ByVal NewGroup As Outlook.OutlookBarGroup) 
 Dim myFolder As Outlook.Folder 
 Set myFolder = myOlApp.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar) 
 NewGroup.Shortcuts.Add myFolder, "Calendar" 
End Sub
```


## See also


#### Concepts


[OutlookBarGroups Object](outlookbargroups-object-outlook.md)

