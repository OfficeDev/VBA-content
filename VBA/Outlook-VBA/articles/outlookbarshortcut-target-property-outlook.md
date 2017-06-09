---
title: OutlookBarShortcut.Target Property (Outlook)
keywords: vbaol11.chm343
f1_keywords:
- vbaol11.chm343
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcut.Target
ms.assetid: 990671c0-bfc5-6b09-26a1-1cdf9d0e143b
ms.date: 06/08/2017
---


# OutlookBarShortcut.Target Property (Outlook)

Returns a  **Variant** indicating the target of the specified shortcut in a **Shortcuts** pane group. Read-only.


## Syntax

 _expression_ . **Target**

 _expression_ A variable that represents an **OutlookBarShortcut** object.


## Remarks

The return type depends on the shortcut type. If the shortcut represents an Outlook folder, the return type is  **Folder** . If the shortcut represents a file-system folder, the return type is an **Object** . If the shortcut represents a file-system path or URL, the return type is a **String** .


## Example

This Microsoft Visual Basic for Applications (VBA) example steps through the shortcuts in the first  **Shortcuts** pane group. It counts the number of shortcuts that are Outlook folders and displays the count.


```vb
Sub DeleteShortcuts() 
 Dim myOlBar As Outlook.OutlookBarPane 
 Dim myolGroup As Outlook.OutlookBarGroup 
 Dim myOlShortcuts As Outlook.OutlookBarShortcuts 
 Dim myOlShortcut As Outlook.OutlookBarShortcut 
 Dim myTop As Integer 
 Dim x As Integer 
 Dim count As Integer 
 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 Set myolGroup = myOlBar.Contents.Groups.Item(1) 
 Set myOlShortcuts = myolGroup.Shortcuts 
 myTop = myOlShortcuts.Count 
 For x = myTop To 1 Step -1 
 Set myOlShortcut = myOlShortcuts.Item(x) 
 If TypeName(myOlShortcut.Target) = "Folder" Then 
 count = count + 1 
 End If 
 Next x 
 MsgBox ("Number of shortcuts that are Outlook folders:" &; count) 
End Sub
```


## See also


#### Concepts


[OutlookBarShortcut Object](outlookbarshortcut-object-outlook.md)

