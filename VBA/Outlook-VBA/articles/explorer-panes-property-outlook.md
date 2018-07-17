---
title: Explorer.Panes Property (Outlook)
keywords: vbaol11.chm2769
f1_keywords:
- vbaol11.chm2769
ms.prod: outlook
api_name:
- Outlook.Explorer.Panes
ms.assetid: b7ec51bd-c8e0-f31e-1f15-42a7514cb433
ms.date: 06/08/2017
---


# Explorer.Panes Property (Outlook)

Returns a  **[Panes](panes-object-outlook.md)** collection object representing the panes displayed by the specified explorer.


## Syntax

 _expression_ . **Panes**

 _expression_ A variable that represents an **Explorer** object.


## Example

This Microsoft Visual Basic for Applications (VBA) example adds a group named "Marketing" as the second group in the  **Shortcuts** pane.


```vb
Sub AddGroup() 
 Dim myolBar As Outlook.OutlookBarPane 
 
 Set myolBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 myolBar.Contents.Groups.Add "Sales", myolBar.Contents.Groups.Count + 1 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

