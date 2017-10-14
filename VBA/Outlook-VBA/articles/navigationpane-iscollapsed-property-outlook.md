---
title: NavigationPane.IsCollapsed Property (Outlook)
keywords: vbaol11.chm2790
f1_keywords:
- vbaol11.chm2790
ms.prod: outlook
api_name:
- Outlook.NavigationPane.IsCollapsed
ms.assetid: 0297c5d3-4c5f-32a4-49eb-85fe0408db60
ms.date: 06/08/2017
---


# NavigationPane.IsCollapsed Property (Outlook)

Returns or sets a  **Boolean** value that determines whether the Navigation Pane is collapsed. Read/write.


## Syntax

 _expression_ . **IsCollapsed**

 _expression_ A variable that represents a **NavigationPane** object.


## Example

The following Visual Basic for Applications (VBA) example collapses the Navigation Pane after hiding all of the modules contained by it.


```vb
Sub CollapseAndHideAllModules() 
 
 Dim objPane As NavigationPane 
 
 
 
 ' Get the NavigationPane object for the 
 
 ' currently displayed Explorer object. 
 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
 
 
 ' Set the DisplayedModuleCount property to 
 
 ' hide all modules contained by the 
 
 ' Navigation Pane. 
 
 objPane.DisplayedModuleCount = 0 
 
 
 
 ' Set the IsCollapsed property to 
 
 ' collapse the Navigation Pane. 
 
 objPane.IsCollapsed = True 
 
 
 
End Sub
```


## See also


#### Concepts


[NavigationPane Object](navigationpane-object-outlook.md)

