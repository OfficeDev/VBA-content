---
title: TextBox.TopMargin Property (Access)
keywords: vbaac10.chm11139
f1_keywords:
- vbaac10.chm11139
ms.prod: access
api_name:
- Access.TextBox.TopMargin
ms.assetid: cd56b2b2-8bb5-b3cf-bacf-13d311e5479b
ms.date: 06/08/2017
---


# TextBox.TopMargin Property (Access)

Along with the  **LeftMargin**, **RightMargin**, and **BottomMargin** properties, specifies the location of information displayed within a text box control. Read/write **Integer**.


## Syntax

 _expression_. **TopMargin**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

A control's displayed information location is measured from the control's left, top, right, or bottom border to the left, top, right, or bottom edge of the displayed information. Setting the  **LeftMargin** or **TopMargin** property to 0 places the displayed information's edge at the very left or top of the control. To use a unit of measurement different from the setting in the regional settings of Windows, specify the unit (for example, cm or in).

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

