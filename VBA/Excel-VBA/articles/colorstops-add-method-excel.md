---
title: ColorStops.Add Method (Excel)
keywords: vbaxl10.chm853074
f1_keywords:
- vbaxl10.chm853074
ms.prod: excel
api_name:
- Excel.ColorStops.Add
ms.assetid: 121c48bf-0b68-89c9-6a03-f7a403b52fee
ms.date: 06/08/2017
---


# ColorStops.Add Method (Excel)

Adds a  **[ColorStop](colorstop-object-excel.md)** object to the specified collection.


## Syntax

 _expression_ . **Add**( **_Position_** )

 _expression_ An expression that returns a **[ColorStops](colorstops-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Position_|Required| **Double**|Represents the position in which to apply the  **ColorStop** .|

### Return Value

ColorStop


## Example

Adds a ColorStop for the active selection.


```vb
Range("A1:A10").Select 
With Selection.Interior.Gradient.ColorStop.Add(1) 
 .ThemeColor = xlThemeColorAccent1 
 .TintAndShade = 0 
End With
```


## See also


#### Concepts


[ColorStops Object](colorstops-object-excel.md)

