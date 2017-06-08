---
title: OptionButton.LabelY Property (Access)
keywords: vbaac10.chm10601
f1_keywords:
- vbaac10.chm10601
ms.prod: access
api_name:
- Access.OptionButton.LabelY
ms.assetid: e5fcac2e-efa7-362f-176f-90ddc53db695
ms.date: 06/08/2017
---


# OptionButton.LabelY Property (Access)

The  **LabelY** property (along with the **LabelX** property) specifies the placement of the label for a new control. Read/write **Integer**.


## Syntax

 _expression_. **LabelY**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks

If the orientation is left to right for a form or report,  **LabelX** and **LabelY** behavior matches standard Microsoft Access left-to-right orientation. For more information about orientation, see the **Orientation** property.

If orientation is right to left, the origin of the coordinate system for  **LabelX** and **LabelY** is the upper right corner of the attached control. A negative number for **LabelX** places the label to the right of the control. A negative number for **LabelY** places the label above the control.

For General and Right alignment when orientation is RTL,  **LabelX** and **LabelY** specify the location of the upper-right corner of the label relative to the upper-right corner of the label's attached control. For Left and Center alignment, **LabelX** and **LabelY** specify the location of the upper-left corner and top center, respectively, of the label relative to the upper-right corner of the label's attached control.


## See also


#### Concepts


[OptionButton Object](optionbutton-object-access.md)

