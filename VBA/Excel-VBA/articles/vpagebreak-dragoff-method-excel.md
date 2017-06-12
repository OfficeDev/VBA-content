---
title: VPageBreak.DragOff Method (Excel)
keywords: vbaxl10.chm156075
f1_keywords:
- vbaxl10.chm156075
ms.prod: excel
api_name:
- Excel.VPageBreak.DragOff
ms.assetid: 93e169e8-e2d6-4cca-bd82-2d11fdc1ae4c
ms.date: 06/08/2017
---


# VPageBreak.DragOff Method (Excel)

Drags a page break out of the print area.


## Syntax

 _expression_ . **DragOff**( **_Direction_** , **_RegionIndex_** )

 _expression_ A variable that represents a **VPageBreak** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[XlDirection](xldirection-enumeration-excel.md)**|The direction in which the page break is dragged.|
| _RegionIndex_|Required| **Long**|The print-area region index for the page break (the region where the mouse pointer is located when the mouse button is pressed if the user drags the page break). If the print area is contiguous, there?s only one print region. If the print area is discontiguous, there?s more than one print region.|

## Remarks

This method exists primarily for the macro recorder. You can use the  **[Delete](vpagebreak-delete-method-excel.md)** method to delete a page break in Visual Basic.


## Example

This example deletes vertical page break one from the active sheet by dragging it off the right edge of print region one.


```vb
ActiveSheet.VPageBreaks(1).DragOff xlToRight, 1
```


## See also


#### Concepts


[VPageBreak Object](vpagebreak-object-excel.md)

