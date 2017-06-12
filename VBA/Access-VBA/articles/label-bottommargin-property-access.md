---
title: Label.BottomMargin Property (Access)
keywords: vbaac10.chm10238
f1_keywords:
- vbaac10.chm10238
ms.prod: access
api_name:
- Access.Label.BottomMargin
ms.assetid: 0d2a1de9-0aea-5bbd-22b7-5b99678240be
ms.date: 06/08/2017
---


# Label.BottomMargin Property (Access)

Along with the  **LeftMargin**, **RightMargin**, and **TopMargin** properties, specifies the location of information displayed within a label control. Read/write **Integer**.


## Syntax

 _expression_. **BottomMargin**

 _expression_ A variable that represents a **Label** object.


## Remarks

A control's displayed information location is the distance measured from the control's left, top, right, or bottom border to the left, top, right, or bottom edge of the displayed information. To use a unit of measurement different from the setting in the regional settings of Windows, specify the unit (for example, cm or in).

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips.


## Example

The following example offsets the caption in the label "EmployeeID_Label" of the "Purchase Orders" form by 100 twips from the bottom of the label's border.


```vb
With Forms.Item("Purchase Orders").Controls.Item("EmployeeID_Label") 
 .BottomMargin = 100 
End With
```


## See also


#### Concepts


[Label Object](label-object-access.md)

