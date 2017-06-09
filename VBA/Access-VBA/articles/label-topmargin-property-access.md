---
title: Label.TopMargin Property (Access)
keywords: vbaac10.chm10235
f1_keywords:
- vbaac10.chm10235
ms.prod: access
api_name:
- Access.Label.TopMargin
ms.assetid: 95432167-4b75-ba84-a75d-57ad3cab35b9
ms.date: 06/08/2017
---


# Label.TopMargin Property (Access)

Along with the  **LeftMargin**, **RightMargin**, and **BottomMargin** properties, specifies the location of information displayed within a label control. Read/write **Integer**.


## Syntax

 _expression_. **TopMargin**

 _expression_ A variable that represents a **Label** object.


## Remarks

A control's displayed information location is measured from the control's left, top, right, or bottom border to the left, top, right, or bottom edge of the displayed information. Setting the  **LeftMargin** or **TopMargin** property to 0 places the displayed information's edge at the very left or top of the control. To use a unit of measurement different from the setting in the regional settings of Windows, specify the unit (for example, cm or in).

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips.


## Example

The following example offsets the caption in the label "EmployeeID_Label" in the "Purchase Orders" form by 100 twips from the top of the label's border.


```vb
With Forms.Item("Purchase Orders").Controls.Item("EmployeeID_Label") 
 .TopMargin = 100 
End With
```


## See also


#### Concepts


[Label Object](label-object-access.md)

