---
title: Style.AddIndent Property (Excel)
keywords: vbaxl10.chm177073
f1_keywords:
- vbaxl10.chm177073
ms.prod: excel
api_name:
- Excel.Style.AddIndent
ms.assetid: 76b9c820-8c94-3cf6-7267-6d2710f07b74
ms.date: 06/08/2017
---


# Style.AddIndent Property (Excel)

Returns or sets a  **Boolean** value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.)


## Syntax

 _expression_ . **AddIndent**

 _expression_ A variable that represents a **Style** object.


## Remarks

Set the value of this property to  **True** to autmatically indent text when the text alignment in the cell is set, either horizontally or vertically, to equal distribution.

To set text alignment to equal distribution, you can set the  **[VerticalAlignment](range-verticalalignment-property-excel.md)** property to **xlVAlignDistributed** when the value of the **[Orientation](range-orientation-property-excel.md)** property is **xlVertical** , and you can set the **[HorizontalAlignment](range-horizontalalignment-property-excel.md)** property to **xlHAlignDistributed** when the value of the **Orientation** property is **xlHorizontal** .


## See also


#### Concepts


[Style Object](style-object-excel.md)

