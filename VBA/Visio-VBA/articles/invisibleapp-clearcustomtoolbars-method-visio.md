---
title: InvisibleApp.ClearCustomToolbars Method (Visio)
keywords: vis_sdr.chm17516115
f1_keywords:
- vis_sdr.chm17516115
ms.prod: visio
api_name:
- Visio.InvisibleApp.ClearCustomToolbars
ms.assetid: 3020ea80-ea8b-3670-865b-329326835a7f
ms.date: 06/08/2017
---


# InvisibleApp.ClearCustomToolbars Method (Visio)

Restores the built-in Microsoft Visio user interface.


## Syntax

 _expression_ . **ClearCustomToolbars**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Nothing


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Calling the  **ClearCustomToolbars** method on an object without custom toolbars has no effect.


## Example

This example shows how to clear custom toolbars for the  **ThisDocument** and **Application** objects and restore the built-in Microsoft Visio toolbars.


```vb
 
Public Sub ClearCustomToolbars_Example() 
 
 'Tell Visio to use the built-in toolbars. 
 ThisDocument.ClearCustomToolbars 
 Visio.Application.ClearCustomToolbars 
 
End Sub
```


