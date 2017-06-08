---
title: OutlookBarShortcuts.BeforeShortcutRemove Event (Outlook)
keywords: vbaol11.chm379
f1_keywords:
- vbaol11.chm379
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcuts.BeforeShortcutRemove
ms.assetid: 4a4107ce-db02-f698-ffae-5a2a4571089c
ms.date: 06/08/2017
---


# OutlookBarShortcuts.BeforeShortcutRemove Event (Outlook)

Occurs before a new shortcut is removed from a group in the  **Shortcuts** pane, either as a result of user action or through program code.


## Syntax

 _expression_ . **BeforeShortcutRemove**( **_Shortcut_** , **_Cancel_** )

 _expression_ A variable that represents an **OutlookBarShortcuts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shortcut_|Required| **[OutlookBarShortcut](outlookbarshortcut-object-outlook.md)**|The  **OutlookBarShortcut** that is being removed.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the shortcut is not removed from the group.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

The following Microsoft Visual Basic for Applications (VBA) example prevents a user from removing a shortcut from the  **Shortcuts** pane. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents myOlShortcuts As Outlook.OutlookBarShortcuts 
Dim myOlBar As Outlook.OutlookBarPane 
 
Sub Initialize_handler() 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 Set myOlShortcuts = myOlBar.Contents.Groups.Item(1).Shortcuts 
End Sub 
 
Private Sub myOlShortcuts_BeforeShortcutRemove(ByVal Shortcut As OutlookBarShortcut, Cancel As Boolean) 
 MsgBox "You are not allowed to remove a shortcut from this group." 
 Cancel = True 
End Sub
```


## See also


#### Concepts


[OutlookBarShortcuts Object](outlookbarshortcuts-object-outlook.md)

