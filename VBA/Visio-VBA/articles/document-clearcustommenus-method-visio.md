---
title: Document.ClearCustomMenus Method (Visio)
keywords: vis_sdr.chm10516110
f1_keywords:
- vis_sdr.chm10516110
ms.prod: visio
api_name:
- Visio.Document.ClearCustomMenus
ms.assetid: 5be16274-151b-e139-8607-76fdb05a4235
ms.date: 06/08/2017
---


# Document.ClearCustomMenus Method (Visio)

Restores the built-in Microsoft Visio user interface.


## Syntax

 _expression_ . **ClearCustomMenus**

 _expression_ A variable that represents a **Document** object.


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


