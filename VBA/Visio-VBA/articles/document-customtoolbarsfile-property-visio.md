---
title: Document.CustomToolbarsFile Property (Visio)
keywords: vis_sdr.chm10513360
f1_keywords:
- vis_sdr.chm10513360
ms.prod: visio
api_name:
- Visio.Document.CustomToolbarsFile
ms.assetid: 1385e027-0cc9-4f3b-a044-ff5731325b25
ms.date: 06/08/2017
---


# Document.CustomToolbarsFile Property (Visio)

Returns or sets the name of the file that defines custom toolbars and status bars for a  **Document** object. Read/write.


## Syntax

 _expression_ . **CustomToolbarsFile**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If the object is not using custom toolbars, the  **CustomToolbarsFile** property returns **Nothing** .


