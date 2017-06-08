---
title: OutlookBarShortcuts.ShortcutAdd Event (Outlook)
keywords: vbaol11.chm377
f1_keywords:
- vbaol11.chm377
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcuts.ShortcutAdd
ms.assetid: d5ddf2ad-0a82-39cb-5bb0-0de389d5c427
ms.date: 06/08/2017
---


# OutlookBarShortcuts.ShortcutAdd Event (Outlook)

Occurs when a new shortcut is added to a  **Shortcuts** pane group.


## Syntax

 _expression_ . **ShortcutAdd**( **_NewShortcut_** )

 _expression_ A variable that represents an **OutlookBarShortcuts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewShortcut_|Required| **[OutlookBarShortcut](outlookbarshortcut-object-outlook.md)**|The shortcut that is being added.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example changes the name of a  **Calendar** shortcut when it is added to the first group in the **Shortcuts pane**. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents myOlSCuts As Outlook.OutlookBarShortcuts 
Dim myOlBar As Outlook.OutlookBarPane 
 
Sub Initialize_handler() 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 Set myOlSCuts = myOlBar.Contents.Groups.Item(1).Shortcuts 
End Sub 
 
Private Sub myOlSCuts_ShortcutAdd(ByVal NewShortcut As outlook.OutlookBarShortcut) 
 Dim myNS As Outlook.NameSpace 
 
 Set myNS = Application.GetNamespace("MAPI") 
 If NewShortcut.Target.Name = "Calendar" Then 
 NewShortcut.Name = myNS.CurrentUser &; "'s Schedules" 
 End If 
End Sub
```


## See also


#### Concepts


[OutlookBarShortcuts Object](outlookbarshortcuts-object-outlook.md)

