---
title: Label.LeftMargin Property (Access)
keywords: vbaac10.chm10234
f1_keywords:
- vbaac10.chm10234
ms.prod: access
api_name:
- Access.Label.LeftMargin
ms.assetid: 7eca4de7-fad8-19f5-c3d2-115cd617755d
ms.date: 06/08/2017
---


# Label.LeftMargin Property (Access)

Along with the  **TopMargin**, **RightMargin**, and **BottomMargin** properties, specifies the location of information displayed within a label control. Read/write **Integer**. .


## Syntax

 _expression_. **LeftMargin**

 _expression_ A variable that represents a **Label** object.


## Remarks

A control's displayed information location is measured from the control's left, top, right, or bottom border to the left, top, right, or bottom edge of the displayed information. Setting the  **LeftMargin** or **TopMargin** property to 0 places the displayed information's edge at the very left or top of the control. To use a unit of measurement different from the setting in the regional settings of Windows, specify the unit (for example, cm or in).

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips.


## See also


#### Concepts


[Label Object](label-object-access.md)

