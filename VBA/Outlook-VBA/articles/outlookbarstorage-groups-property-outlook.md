---
title: OutlookBarStorage.Groups Property (Outlook)
keywords: vbaol11.chm372
f1_keywords:
- vbaol11.chm372
ms.prod: outlook
api_name:
- Outlook.OutlookBarStorage.Groups
ms.assetid: 9b324d3d-3ab6-1e24-962f-19812b6b8ed0
ms.date: 06/08/2017
---


# OutlookBarStorage.Groups Property (Outlook)

Returns an  **[OutlookBarGroups](outlookbargroups-object-outlook.md)** object representing the set of groups in the **Shortcuts** pane. Read-only.


## Syntax

 _expression_ . **Groups**

 _expression_ A variable that represents an **OutlookBarStorage** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the number of groups in the  **Shortcuts** pane.


```vb
Sub CountOlBarGroups()     Dim myOlBar As Outlook.OutlookBarPane     Dim myCount As Integer      Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar")     myCount = myOlBar.Contents.Groups.Count     MsgBox "There are " &; myCount &; " groups in the Shortcuts pane" End Sub
```


## See also


#### Concepts


[OutlookBarStorage Object](outlookbarstorage-object-outlook.md)

