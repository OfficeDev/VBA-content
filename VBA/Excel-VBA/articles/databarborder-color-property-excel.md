---
title: DataBarBorder.Color Property (Excel)
keywords: vbaxl10.chm885074
f1_keywords:
- vbaxl10.chm885074
ms.prod: excel
api_name:
- Excel.DataBarBorder.Color
ms.assetid: a16439a9-c086-9c42-8496-9a16d9011689
ms.date: 06/08/2017
---


# DataBarBorder.Color Property (Excel)

Returns an object that specifies the color of the border of data bars specified by a conditional formatting rule. Read-only


## Syntax

 _expression_ . **Color**

 _expression_ A variable that represents a **[DataBarBorder](databarborder-object-excel.md)** object.


### Return Value

 **[FormatColor](formatcolor-object-excel.md)**


## Example

The following code example selects a range of cells and adds a data bar conditional formatting rule to that range. It then uses the  **[BarBorder](databar-barborder-property-excel.md)** property to retrieve the **DataBarBorder** object associated with that rule, and uses the **Color** property of that object to retrieve the **FormatColor** object to set the color and tint of the data bar borders.


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

