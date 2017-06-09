---
title: OutlookBarShortcuts.BeforeShortcutAdd Event (Outlook)
keywords: vbaol11.chm378
f1_keywords:
- vbaol11.chm378
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcuts.BeforeShortcutAdd
ms.assetid: b31d495f-8288-a2ee-1429-6face8281787
ms.date: 06/08/2017
---


# OutlookBarShortcuts.BeforeShortcutAdd Event (Outlook)

Occurs before a new shortcut is added to a group in the  **Shortcuts** pane, either as a result of user action or through program code.


## Syntax

 _expression_ . **BeforeShortcutAdd**( **_Cancel_** )

 _expression_ A variable that represents an **OutlookBarShortcuts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the shortcut is not added to the group.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

The following Microsoft Visual Basic for Applications (VBA) example prevents a user from adding a shortcut to the first group in the  **Shortcuts** pane. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents myOlShortcuts As Outlook.OutlookBarShortcuts 
Dim myOlBar As Outlook.OutlookBarPane 
 
Sub Initialize_handler() 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 Set myOlShortcuts = myOlBar.Contents.Groups.Item(1).Shortcuts 
End Sub 
 
Private Sub myOlShortcuts_BeforeShortcutAdd(Cancel As Boolean) 
 MsgBox "You are not allowed to add a shortcut to this group." 
 Cancel = True 
End Sub
```


## See also


#### Concepts


[OutlookBarShortcuts Object](outlookbarshortcuts-object-outlook.md)

