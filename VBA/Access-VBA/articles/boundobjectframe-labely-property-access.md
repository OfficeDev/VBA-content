---
title: BoundObjectFrame.LabelY Property (Access)
keywords: vbaac10.chm10948
f1_keywords:
- vbaac10.chm10948
ms.prod: access
api_name:
- Access.BoundObjectFrame.LabelY
ms.assetid: 1c68016d-9be5-b550-1b97-1840ed36f974
ms.date: 06/08/2017
---


# BoundObjectFrame.LabelY Property (Access)

The  **LabelY** property (along with the **LabelX** property) specifies the placement of the label for a new control. Read/write **Integer**.



## Syntax

 _expression_. **LabelY**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

If the orientation is left to right for a form or report,  **LabelX** and **LabelY** behavior matches standard Microsoft Access left-to-right orientation. For more information about orientation, see the **Orientation** property.

If orientation is right to left, the origin of the coordinate system for  **LabelX** and **LabelY** is the upper right corner of the attached control. A negative number for **LabelX** places the label to the right of the control. A negative number for **LabelY** places the label above the control.

For General and Right alignment when orientation is RTL,  **LabelX** and **LabelY** specify the location of the upper-right corner of the label relative to the upper-right corner of the label's attached control. For Left and Center alignment, **LabelX** and **LabelY** specify the location of the upper-left corner and top center, respectively, of the label relative to the upper-right corner of the label's attached control.


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

