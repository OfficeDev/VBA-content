---
title: TextBox.SelectionMargin Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: afa418ab-4da0-df67-5545-dc4633e057e4
ms.date: 06/08/2017
---


# TextBox.SelectionMargin Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether the user can select a line of text by clicking in the region to the left of the text. Read/write.


## Syntax

 _expression_. **SelectionMargin**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

 **True** if clicking in margin causes selection of text (default), **False** if clicking in margin does not cause selection of text.

When the  **SelectionMargin** property is **True**, the selection margin occupies a thin strip along the left edge of a control's edit region. When set to  **False**, the entire edit region can store text.

If the  **SelectionMargin** property is set to **True** when a control is printed, the selection margin also prints.


