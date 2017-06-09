---
title: Explorer.NavigationPane Property (Outlook)
keywords: vbaol11.chm2782
f1_keywords:
- vbaol11.chm2782
ms.prod: outlook
api_name:
- Outlook.Explorer.NavigationPane
ms.assetid: 9ff92a76-d1cd-e338-2f45-e3e5c79c136e
ms.date: 06/08/2017
---


# Explorer.NavigationPane Property (Outlook)

Returns a  **[NavigationPane](navigationpane-object-outlook.md)** object that represents the Navigation Pane for an **[Explorer](explorer-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **NavigationPane**

 _expression_ A variable that represents an **Explorer** object.


## Remarks

Some  **Explorer** objects may not have an associated **NavigationPane** object. In such cases, this property returns **Null** ( **Nothing** in Visual Basic.)


## Example

The following Visual Basic for Applications (VBA) sample retrieves the  **NavigationPane** object from the active **Explorer** object and then displays information about the number of navigation modules contained and displayed by the object.


```vb
Sub DisplayModuleCounts() 
 
 Dim objPane As NavigationPane 
 
 
 
 ' Get the NavigationPane object for the 
 
 ' currently displayed Explorer object. 
 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
 
 
 ' Display information about modules contained 
 
 ' by the NavigationPane object. 
 
 MsgBox "The Navigation Pane currently contains " &; _ 
 
 objPane.Modules.Count &; _ 
 
 " modules, of which " &; _ 
 
 objPane.DisplayedModuleCount &; _ 
 
 " are displayed." 
 
 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

