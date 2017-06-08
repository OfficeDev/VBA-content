---
title: ComboBox.TextAlign Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: e80b00a9-2020-3769-0d0d-84e66273a1ce
ms.date: 06/08/2017
---


# ComboBox.TextAlign Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies how text is aligned in a control. Read/write.


## Syntax

 _expression_. **TextAlign**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The settings for  **TextAlign** are:



|**Value**|**Description**|
|:-----|:-----|
|1|Aligns the first character of displayed text with the left edge of the control's display or edit area (default).|
|2|Centers the text in the control's display or edit area.|
|3|Aligns the last character of displayed text with the right edge of the control's display or edit area.|
The  **TextAlign** property only affects the edit region; this property has no effect on the alignment of text in the list.


