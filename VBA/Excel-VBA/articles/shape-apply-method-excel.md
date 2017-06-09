---
title: Shape.Apply Method (Excel)
keywords: vbaxl10.chm636074
f1_keywords:
- vbaxl10.chm636074
ms.prod: excel
api_name:
- Excel.Shape.Apply
ms.assetid: fe094baf-76d7-8418-aa34-c90d37f95def
ms.date: 06/08/2017
---


# Shape.Apply Method (Excel)

Applies to the specified shape formatting that's been copied by using the  **[PickUp](shape-pickup-method-excel.md)** method.


## Syntax

 _expression_ . **Apply**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-excel.md)

