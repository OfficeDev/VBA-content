---
title: ControlFormat.RemoveAllItems Method (Excel)
keywords: vbaxl10.chm630074
f1_keywords:
- vbaxl10.chm630074
ms.prod: excel
api_name:
- Excel.ControlFormat.RemoveAllItems
ms.assetid: de8e1721-45e1-eca9-d35d-7d72c32dc0bf
ms.date: 06/08/2017
---


# ControlFormat.RemoveAllItems Method (Excel)

Removes all entries from a Microsoft Excel list box or combo box.


## Syntax

 _expression_ . **RemoveAllItems**

 _expression_ A variable that represents a **ControlFormat** object.


## Example

This example removes all items from a list box. If  `Shapes(2)` doesn't represent a list box, this example fails.


```vb
Worksheets(1).Shapes(2).ControlFormat.RemoveAllItems
```


## See also


#### Concepts


[ControlFormat Object](controlformat-object-excel.md)

