---
title: Application.ViewEditCombination Method (Project)
keywords: vbapj.chm304
f1_keywords:
- vbapj.chm304
ms.prod: project-server
api_name:
- Project.Application.ViewEditCombination
ms.assetid: f5d49a1d-7ead-e704-7be2-8d06e54e221f
ms.date: 06/08/2017
---


# Application.ViewEditCombination Method (Project)

Creates, edits, or copies a combination view.


## Syntax

 _expression_. **ViewEditCombination**( ** _Name_**, ** _Create_**, ** _NewName_**, ** _TopView_**, ** _BottomView_**, ** _ShowInMenu_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a two-pane view to edit, create, or copy. The default is the name of the active view.|
| _Create_|Optional|**Boolean**|**True** if Project creates a two-pane view. If NewName is an empty string (""), the new view is given the name specified with Name. Otherwise, the new view is a copy of the view specified with Name and is given the name specified with NewName. The default value is **False.**|
| _NewName_|Optional|**String**|A new name for the view specified with Name (Create is  **False** ) or a name for the new view just created (Create is **True** ). If NewName is an empty string and Create is **False**, the view specified with Name retains its current name. The default value is **False.**|
| _TopView_|Optional|**String**|The name of the view to display in the upper pane. The view specified by Name displays in the lower pane.|
| _BottomView_|Optional|**String**|The name of the view to display in the lower pane. The view specified by Name displays in the upper pane.|
| _ShowInMenu_|Optional|**Boolean**|**True** if the view name appears on the **Other Views** drop-down menu. The default value is **False.**|

### Return Value

 **Boolean**


## Example

The following example creates a combination view with the Resource Sheet in the upper pane and the Resource Graph in the lower pane.


```vb
Sub CheckResourcesView() 
 ViewEditCombination Name:="Check Resources View", Create:=True, _ 
 TopView:="Resource Sheet", BottomView:="Resource Graph" 
End Sub
```


