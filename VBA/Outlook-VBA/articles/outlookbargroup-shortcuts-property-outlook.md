---
title: OutlookBarGroup.Shortcuts Property (Outlook)
keywords: vbaol11.chm326
f1_keywords:
- vbaol11.chm326
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroup.Shortcuts
ms.assetid: a6a5031e-4ca2-4b4f-00b3-298af2361cec
ms.date: 06/08/2017
---


# OutlookBarGroup.Shortcuts Property (Outlook)

Returns an  **[OutlookBarShortcuts](outlookbarshortcuts-object-outlook.md)** collection of shortcuts contained within the **Shortcuts** pane. Read-only.


## Syntax

 _expression_ . **Shortcuts**

 _expression_ A variable that represents an **OutlookBarGroup** object.


## Example

This Microsoft Visual Basic for Applications (VBA) example deletes all empty groups in the  **Shortcuts** pane.


```vb
Sub DeleteEmptyGroups() 
 Dim myOlBar As Outlook.OutlookBarPane 
 Dim myOlGroup As Outlook.OutlookBarGroup 
 Dim x As Integer 
 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 For x = myOlBar.Contents.Groups.Count To 1 Step -1 
 Set myOlGroup = myOlBar.Contents.Groups.Item(x) 
 If myOlGroup.Shortcuts.Count = 0 Then 
 myOlBar.Contents.Groups.Remove x 
 End If 
 Next x 
End Sub
```


## See also


#### Concepts


[OutlookBarGroup Object](outlookbargroup-object-outlook.md)

