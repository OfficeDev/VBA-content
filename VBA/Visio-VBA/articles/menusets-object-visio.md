---
title: MenuSets Object (Visio)
keywords: vis_sdr.chm10175
f1_keywords:
- vis_sdr.chm10175
ms.prod: visio
api_name:
- Visio.MenuSets
ms.assetid: 6a49d679-abdb-2bd4-134b-c61ea3f196e8
ms.date: 06/08/2017
---


# MenuSets Object (Visio)

Includes a  **MenuSet** object for each Microsoft Visio window context that has menus.


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

To retrieve a  **MenuSets** collection, use the **MenuSets** property of a **UIObject** object.

The default property of a  **MenuSets** collection is **Item** .

Unlike other Visio collections, the  **MenuSets** collection is indexed starting with zero (0) rather than 1.

A  **MenuSet** object is identified in the **MenuSets** collection by its **SetID** property, which corresponds to a Visio window context. For a list of **SetID** values for **MenuSet** objects, see the **SetID** property.


