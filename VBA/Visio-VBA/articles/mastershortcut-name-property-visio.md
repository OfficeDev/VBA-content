---
title: MasterShortcut.Name Property (Visio)
keywords: vis_sdr.chm16013930
f1_keywords:
- vis_sdr.chm16013930
ms.prod: visio
api_name:
- Visio.MasterShortcut.Name
ms.assetid: 0c95f270-3810-b924-e55a-dd758b92cacd
ms.date: 06/08/2017
---


# MasterShortcut.Name Property (Visio)

Specifies the name of an object. Read-only.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **MasterShortcut** object.


### Return Value

String


## Remarks




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Name** property to get or set a **Hyperlink** , **Layer** , **Master** , **MasterShortcut** , **Page** , **Shape** , **Style** , or **Row** object's local name. Use the **NameU** property to get or set its universal name.


