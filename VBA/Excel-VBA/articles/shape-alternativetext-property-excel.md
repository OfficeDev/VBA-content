---
title: Shape.AlternativeText Property (Excel)
keywords: vbaxl10.chm636132
f1_keywords:
- vbaxl10.chm636132
ms.prod: excel
api_name:
- Excel.Shape.AlternativeText
ms.assetid: 40b53b31-c4e2-0fd8-1a37-fa1e88ccd2be
ms.date: 06/08/2017
---


# Shape.AlternativeText Property (Excel)

Returns or sets the descriptive (alternative) text string for a  **[Shape](shape-object-excel.md)** object when the object is saved to a Web page. Read/write **String** .


## Syntax

 _expression_ . **AlternativeText**

 _expression_ A variable that represents a **Shape** object.


## Remarks

The alternative text can be displayed either in place of the shape's image in the Web browser , or directly over the image when the mouse pointer hovers over the image (in browsers that support these features).


## Example

This example sets the alternative text for the first shape on the first worksheet to a description of the shape.


```vb
Worksheets(1).Shapes(1).AlternativeText = "Concentric circles"
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

