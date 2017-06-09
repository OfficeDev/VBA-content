---
title: InvisibleApp.CustomToolbarsFile Property (Visio)
keywords: vis_sdr.chm17513360
f1_keywords:
- vis_sdr.chm17513360
ms.prod: visio
api_name:
- Visio.InvisibleApp.CustomToolbarsFile
ms.assetid: 0874023f-1e61-7842-be7d-9abe5c4ec63c
ms.date: 06/08/2017
---


# InvisibleApp.CustomToolbarsFile Property (Visio)

Returns or sets the name of the file that defines custom toolbars and status bars for an  **InvisibleApp** object. Read/write.


## Syntax

 _expression_ . **CustomToolbarsFile**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If the object is not using custom toolbars, the  **CustomToolbarsFile** property returns **Nothing** .


