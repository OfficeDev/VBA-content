---
title: ComboBox.LabelX Property (Access)
keywords: vbaac10.chm11430
f1_keywords:
- vbaac10.chm11430
ms.prod: access
api_name:
- Access.ComboBox.LabelX
ms.assetid: 3878f4b3-6f0d-a857-1988-9fae59c0302b
ms.date: 06/08/2017
---


# ComboBox.LabelX Property (Access)

The  **LabelX** property (along with the **LabelY** property) specifies the placement of the label for a new control. Read/write **Integer**.



## Syntax

 _expression_. **LabelX**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

If the orientation is left to right for a form or report,  **LabelX** and **LabelY** behavior matches standard Microsoft Access left-to-right orientation. For more information about orientation, see the **Orientation** property.

If orientation is right to left, the origin of the coordinate system for  **LabelX** and **LabelY** is the upper right corner of the attached control. A negative number for **LabelX** places the label to the right of the control. A negative number for **LabelY** places the label above the control.

For General and Right alignment when orientation is RTL,  **LabelX** and **LabelY** specify the location of the upper-right corner of the label relative to the upper-right corner of the label's attached control. For Left and Center alignment, **LabelX** and **LabelY** specify the location of the upper-left corner and top center, respectively, of the label relative to the upper-right corner of the label's attached control.


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

