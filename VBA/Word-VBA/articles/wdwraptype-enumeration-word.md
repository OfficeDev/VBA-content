---
title: WdWrapType Enumeration (Word)
ms.prod: word
api_name:
- Word.WdWrapType
ms.assetid: b572e6e5-707e-be2f-afd5-158369ed6e8e
ms.date: 06/08/2017
---


# WdWrapType Enumeration (Word)

Specifies how to wrap text around a shape.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdWrapInline**|7|Places shapes in line with text.|
| **wdWrapNone**|3|Places shape in front of text. See also  **wdWrapFront** .|
| **wdWrapSquare**|0|Wraps text around the shape. Line continuation is on the oposite side of the shape.|
| **wdWrapThrough**|2|Wraps text around the shape.|
| **wdWrapTight**|1|Wraps text close to the shape.|
| **wdWrapTopBottom**|4|Places text above and below the shape.|
| **wdWrapBehind**|5|Places shape behind text.|
| **wdWrapFront**|6|Places shape in front of text.|

## Remarks

Used with the  **Type** property of the **WrapFormat** object. You can view what each wrap type looks like on the **Text Wrapping** tab of the **Advanced Layout** dialog box.


 **Note**   **wdWrapSquare** , **wdWrapTight** , and **wdWrapThrough** all function in essentially the same way.


