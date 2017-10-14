---
title: Application.LayoutSelectionNow Method (Project)
keywords: vbapj.chm2399
f1_keywords:
- vbapj.chm2399
ms.prod: project-server
api_name:
- Project.Application.LayoutSelectionNow
ms.assetid: 79d8521a-2760-7e73-f430-f39dc7747cd8
ms.date: 06/08/2017
---


# Application.LayoutSelectionNow Method (Project)

Positions the selected task boxes in the active Network Diagram view according to its layout options.


## Syntax

 _expression_. **LayoutSelectionNow**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

Layout options can be set with the  **BoxLayout** and **BoxLinks** methods.

The  **LayoutSelectionNow** method is only available when a Network Diagram view is active.


## Example

The following example positions the selected boxes in a top-down layout.


```vb
Sub Layout_SelectionNow() 
 
 'Activate Network Diagram view 
 ViewApply Name:="Network &;Diagram" 
 
 BoxSet Action:=pjBoxAddToSelection, TaskID:=2 
 BoxLayout LayoutMode:=pjLayoutManual, LayoutScheme:=pjLayoutTopDownByDay 
 
 LayoutSelectionNow 
End Sub
```


