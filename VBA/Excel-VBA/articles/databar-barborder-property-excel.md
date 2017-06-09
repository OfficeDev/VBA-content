---
title: Databar.BarBorder Property (Excel)
keywords: vbaxl10.chm810094
f1_keywords:
- vbaxl10.chm810094
ms.prod: excel
api_name:
- Excel.Databar.BarBorder
ms.assetid: d573e56e-cd02-c67e-ace8-8e8bdf2efd00
ms.date: 06/08/2017
---


# Databar.BarBorder Property (Excel)

Returns an object that specifies the border of a data bar. Read-only


## Syntax

 _expression_ . **BarBorder**

 _expression_ A variable that represents a **[Databar](databar-object-excel.md)** object.


### Return Value

 **[DataBarBorder](databarborder-object-excel.md)**


## Example

The following code example selects a range of cells, adds a data bar conditional formatting rule to that range, uses the  **BarBorder** property to retrieve the **DataBarBorder** object associated with that rule, and then sets the data bar's color, tint, and type.


```vb
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
With myDataBar.BarBorder 
 .Type = xlDataBarBorderSolid 
 .Color.ThemeColor = xlThemeColorAccent2 
 .Color.TintAndShade = 0 
End With 

```


## See also


#### Concepts


[Databar Object](databar-object-excel.md)

