---
title: Interior.PatternColorIndex Property (Excel)
keywords: vbaxl10.chm551078
f1_keywords:
- vbaxl10.chm551078
ms.prod: excel
api_name:
- Excel.Interior.PatternColorIndex
ms.assetid: e7e89281-e179-bea9-58bf-110f7a4aab8d
ms.date: 06/08/2017
---


# Interior.PatternColorIndex Property (Excel)

Returns or sets the color of the interior pattern as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-excel.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Read/write **Long** .


## Syntax

 _expression_ . **PatternColorIndex**

 _expression_ A variable that represents an **Interior** object.


## Remarks

Set this property to  **xlColorIndexAutomatic** to specify the automatic pattern for cells or the automatic fill style for drawing objects. Set this property to **xlColorIndexNone** to specify that you don't want a pattern (this is the same as setting the **Pattern** property of the **Interior** object to **xlPatternNone** ).


## Example

This example sets the color of the interior pattern for rectangle one on Sheet1.


```vb
With Worksheets("Sheet1").Rectangles(1).Interior 
 .Pattern = xlChecker 
 .PatternColorIndex = 5 
End With
```


## See also


#### Concepts


[Interior Object](interior-object-excel.md)

