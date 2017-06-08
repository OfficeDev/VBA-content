---
title: ComboBox.ScrollBarAlign Property (Access)
keywords: vbaac10.chm11466
f1_keywords:
- vbaac10.chm11466
ms.prod: access
api_name:
- Access.ComboBox.ScrollBarAlign
ms.assetid: ded4533c-2879-d57f-b6ff-cccd20a88090
ms.date: 06/08/2017
---


# ComboBox.ScrollBarAlign Property (Access)

You can use the  **ScrollBarAlign** to specify or determine the alignment of a vertical scroll bar. Read/write **Byte**.


## Syntax

 _expression_. **ScrollBarAlign**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **ScrollBarAlign** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|System|0|Vertical scroll bar is placed on the left if the form or report  **Orientation** property is right to left; and on the right if the form or report **Orientation** property is left to right.|
|Right|1|Aligns vertical scroll bar on the right side of the control.|
|Left|2|Aligns vertical scroll bar on the left side of the control.|
For combo and list boxes,  **ScrollBarAlign** also controls the placement of the box button above the scroll bar.


## Example

The following example aligns the vertical scroll bar on the left side of the "Country" combo box in the "International Shipping" form.


```vb
Forms("International Shipping").Controls("Country").ScrollBarAlign = 2
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

