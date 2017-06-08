---
title: OutlookBarGroups.BeforeGroupAdd Event (Outlook)
keywords: vbaol11.chm356
f1_keywords:
- vbaol11.chm356
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroups.BeforeGroupAdd
ms.assetid: 7bce246a-69fa-0dcd-4c43-fbfc43385864
ms.date: 06/08/2017
---


# OutlookBarGroups.BeforeGroupAdd Event (Outlook)

Occurs before a new group is added to the  **Shortcuts** pane, either as a result of user action or through program code.


## Syntax

 _expression_ . **BeforeGroupAdd**( **_Cancel_** )

 _expression_ A variable that represents an **OutlookBarGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the group is not added to the **Shortcuts** pane.|

## Remarks

 This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Visual Basic for Applications (VBA) example prevents the user from adding a group to the  **Shortcuts** pane. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `Initialize_handler` routine must be called before the event procedure can be called by Outlook.


```vb
Dim WithEvents myOlGroups As Outlook.OutlookBarGroups 
Dim myOlBar As Outlook.OutlookBarPane 
 
Sub Initialize_handler() 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 Set myOlGroups = myOlBar.Contents.Groups 
End Sub 
 
Private Sub myOlGroups_BeforeGroupAdd(Cancel As Boolean) 
 Cancel = True 
End Sub
```


## See also


#### Concepts


[OutlookBarGroups Object](outlookbargroups-object-outlook.md)

