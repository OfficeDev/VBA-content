---
title: NavigationControl.ScrollBarAlign Property (Access)
keywords: vbaac10.chm11133
f1_keywords:
- vbaac10.chm11133
ms.prod: access
api_name:
- Access.NavigationControl.ScrollBarAlign
ms.assetid: b685e196-513e-fe57-d993-d1e2f4051a4c
ms.date: 06/08/2017
---


# NavigationControl.ScrollBarAlign Property (Access)

You can use the  **ScrollBarAlign** to specify or determine the alignment of a vertical scroll bar. Read/write **Byte**.


## Syntax

 _expression_. **ScrollBarAlign**

 _expression_ A variable that represents a **NavigationControl** object.


## Remarks

The  **ScrollBarAlign** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|System|0|Vertical scroll bar is placed on the left if the form or report  **Orientation** property is right to left; and on the right if the form or report **Orientation** property is left to right.|
|Right|1|Aligns vertical scroll bar on the right side of the control.|
|Left|2|Aligns vertical scroll bar on the left side of the control.|

## Example

The following example aligns the vertical scroll bar on the left side of the "Country" combo box in the "International Shipping" form.


```vb
Forms("International Shipping").Controls("Country").ScrollBarAlign = 2
```


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

