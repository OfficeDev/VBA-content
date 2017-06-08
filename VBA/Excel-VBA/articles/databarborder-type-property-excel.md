---
title: DataBarBorder.Type Property (Excel)
keywords: vbaxl10.chm885073
f1_keywords:
- vbaxl10.chm885073
ms.prod: excel
api_name:
- Excel.DataBarBorder.Type
ms.assetid: 9364fadd-5dba-d8a2-a704-a4876173e4a2
ms.date: 06/08/2017
---


# DataBarBorder.Type Property (Excel)

Returns or sets the type of border for data bars specified by a conditional formatting rule. Read/write


## Syntax

 _expression_ . **Type**

 _expression_ A variable that represents a **[DataBarBorder](databarborder-object-excel.md)** object.


### Return Value

 **[XlDataBarBorderType](xldatabarbordertype-enumeration-excel.md)**


## Example

The following code example selects a range of cells, adds a data bar conditional formatting rule to that range, uses the  **[BarBorder](databar-barborder-property-excel.md)** property to retrieve the **DataBarBorder** object associated with that rule, and then sets the data bar color, tint, and type.


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


[DataBarBorder Object](databarborder-object-excel.md)

