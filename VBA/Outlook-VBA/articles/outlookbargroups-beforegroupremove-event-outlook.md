---
title: OutlookBarGroups.BeforeGroupRemove Event (Outlook)
keywords: vbaol11.chm357
f1_keywords:
- vbaol11.chm357
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroups.BeforeGroupRemove
ms.assetid: b3ad5d29-c906-ebe7-02b7-145091ddccce
ms.date: 06/08/2017
---


# OutlookBarGroups.BeforeGroupRemove Event (Outlook)

Occurs before a new group is removed from the  **Shortcuts** pane, either as a result of user action or through program code.


## Syntax

 _expression_ . **BeforeGroupRemove**( **_Group_** , **_Cancel_** )

 _expression_ A variable that represents an **OutlookBarGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Group_|Required| **[OutlookBarGroup](outlookbargroup-object-outlook.md)**|The  **OutlookBarGroup** that is being removed.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the group is not removed from the **Shortcuts** pane.|

## Remarks

 This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example prevents the user from removing a group from the  **Shortcuts** pane. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `Initialize_handler` routine must be called before the event procedure can be called by Outlook. You will still be prompted when you try to delete a shortcut. However, the group will not be deleted even if you clicked **Yes**.


```vb
Dim WithEvents myOlGroups As Outlook.OutlookBarGroups 
Dim myOlBar As Outlook.OutlookBarPane 
 
Sub Initialize_handler() 
 Set myOlBar = Application.ActiveExplorer.Panes.item("OutlookBar") 
 Set myOlGroups = myOlBar.Contents.Groups 
End Sub 
 
Private Sub myOlGroups_BeforeGroupRemove(ByVal Group As OutlookBarGroup, Cancel As Boolean) 
 Cancel = True 
End Sub
```


## See also


#### Concepts


[OutlookBarGroups Object](outlookbargroups-object-outlook.md)

