---
title: ShapeRange.Apply Method (Excel)
keywords: vbaxl10.chm640078
f1_keywords:
- vbaxl10.chm640078
ms.prod: excel
api_name:
- Excel.ShapeRange.Apply
ms.assetid: 34acef44-7075-ffc1-199c-3396e17caafe
ms.date: 06/08/2017
---


# ShapeRange.Apply Method (Excel)

Applies to the specified shape formatting that's been copied by using the  **[PickUp](shaperange-pickup-method-excel.md)** method.


## Syntax

 _expression_ . **Apply**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example copies the formatting of shape one on  `myDocument` and then applies the copied formatting to shape two.


```vb
Set myDocument = Worksheets(1) 
With myDocument 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

