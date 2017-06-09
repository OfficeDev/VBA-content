---
title: Shape.Hyperlink Property (Excel)
keywords: vbaxl10.chm636117
f1_keywords:
- vbaxl10.chm636117
ms.prod: excel
api_name:
- Excel.Shape.Hyperlink
ms.assetid: 97c87fda-91a5-b5db-a82b-6ba1465442fa
ms.date: 06/08/2017
---


# Shape.Hyperlink Property (Excel)

Returns a  **[Hyperlink](hyperlink-object-excel.md)** object that represents the hyperlink for the shape.


## Syntax

 _expression_ . **Hyperlink**

 _expression_ A variable that represents a **Shape** object.


## Example

This example loads the document attached to the hyperlink on shape one.


```vb
Worksheets(1).Shapes(1).Hyperlink.Follow NewWindow:=True
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

