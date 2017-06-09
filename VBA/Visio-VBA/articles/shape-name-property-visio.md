---
title: Shape.Name Property (Visio)
keywords: vis_sdr.chm11213930
f1_keywords:
- vis_sdr.chm11213930
ms.prod: visio
api_name:
- Visio.Shape.Name
ms.assetid: a0708af0-a813-7539-c43f-049009f1ab62
ms.date: 06/08/2017
---


# Shape.Name Property (Visio)

Specifies the name of an object. Read-only.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

A cell has both a local name and a universal name. The local name differs depending on the locale for which the running version of Microsoft Windows is installed. The universal name is the same regardless of what locale is installed. To get the universal name of a cell, use the  **Name** property. To get the local name, use the **LocalName** property.




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Name** property to get or set a **Hyperlink** , **Layer** , **Master** , **MasterShortcut** , **Page** , **Shape** , **Style** , or **Row** object's local name. Use the **NameU** property to get or set its universal name.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVShape.Name**
    

