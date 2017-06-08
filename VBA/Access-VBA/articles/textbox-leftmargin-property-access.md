---
title: TextBox.LeftMargin Property (Access)
keywords: vbaac10.chm11138
f1_keywords:
- vbaac10.chm11138
ms.prod: access
api_name:
- Access.TextBox.LeftMargin
ms.assetid: 9c5b798b-4afe-85be-aa06-eeff98888850
ms.date: 06/08/2017
---


# TextBox.LeftMargin Property (Access)

Along with the  **TopMargin**, **RightMargin**, and **BottomMargin** properties. specifies the location of information displayed within a text box control. Read/write **Integer**. .


## Syntax

 _expression_. **LeftMargin**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

A control's displayed information location is measured from the control's left, top, right, or bottom border to the left, top, right, or bottom edge of the displayed information. Setting the  **LeftMargin** or **TopMargin** property to 0 places the displayed information's edge at the very left or top of the control. To use a unit of measurement different from the setting in the regional settings of Windows, specify the unit (for example, cm or in).

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

