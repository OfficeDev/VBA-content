---
title: Application.CustomToolbarsFile Property (Visio)
keywords: vis_sdr.chm10013360
f1_keywords:
- vis_sdr.chm10013360
ms.prod: visio
api_name:
- Visio.Application.CustomToolbarsFile
ms.assetid: e4759ee0-1128-8238-ad0b-47ad365ce88d
ms.date: 06/08/2017
---


# Application.CustomToolbarsFile Property (Visio)

Returns or sets the name of the file that defines custom toolbars and status bars for an  **Application** object. Read/write.


## Syntax

 _expression_ . **CustomToolbarsFile**

 _expression_ A variable that represents an **Application** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If the object is not using custom toolbars, the  **CustomToolbarsFile** property returns **Nothing** .


