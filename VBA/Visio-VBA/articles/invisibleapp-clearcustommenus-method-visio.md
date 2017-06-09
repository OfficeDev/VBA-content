---
title: InvisibleApp.ClearCustomMenus Method (Visio)
keywords: vis_sdr.chm17516110
f1_keywords:
- vis_sdr.chm17516110
ms.prod: visio
api_name:
- Visio.InvisibleApp.ClearCustomMenus
ms.assetid: bf075a03-4b3f-4f38-a6b4-6fbc2cc95fe1
ms.date: 06/08/2017
---


# InvisibleApp.ClearCustomMenus Method (Visio)

Restores the built-in Microsoft Visio user interface.


## Syntax

 _expression_ . **ClearCustomMenus**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Nothing


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Calling the  **ClearCustomMenus** method on an object without custom menus has no effect.


## Example

This example shows how to clear custom menus for the  **ThisDocument** and **Application** objects and restore the built-in Visio menus.


```vb
 
Public Sub ClearCustomMenus_Example() 
 
 'Tell Visio to use the built-in menus. 
 ThisDocument.ClearCustomMenus 
 Visio.Application.ClearCustomMenus 
 
End Sub
```


