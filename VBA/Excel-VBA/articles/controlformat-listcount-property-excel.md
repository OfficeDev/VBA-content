---
title: ControlFormat.ListCount Property (Excel)
keywords: vbaxl10.chm630081
f1_keywords:
- vbaxl10.chm630081
ms.prod: excel
api_name:
- Excel.ControlFormat.ListCount
ms.assetid: 9f7b60aa-8bf9-a7ec-c198-0a6f6316cc3c
ms.date: 06/08/2017
---


# ControlFormat.ListCount Property (Excel)

Returns the number of entries in a list box or combo box. Returns 0 (zero) if there are no entries in the list. Read-only  **Long** .


## Syntax

 _expression_ . **ListCount**

 _expression_ A variable that represents a **ControlFormat** object.


## Example

This example adjusts a combo box to display all entries in its list. If  `Shapes(1)` does not represent a combo box, this example fails.


```vb
Set cf = Worksheets(1).Shapes(1).ControlFormat 
cf.DropDownLines = cf.ListCount
```


## See also


#### Concepts


[ControlFormat Object](controlformat-object-excel.md)

