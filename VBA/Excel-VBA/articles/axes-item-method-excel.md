---
title: Axes.Item Method (Excel)
keywords: vbaxl10.chm572074
f1_keywords:
- vbaxl10.chm572074
ms.prod: excel
api_name:
- Excel.Axes.Item
ms.assetid: 5e89a576-d2a0-d069-4db6-fc1cf9bd6c61
ms.date: 06/08/2017
---


# Axes.Item Method (Excel)

Returns a single  **[Axis](axis-object-excel.md)** object from an **Axes** collection.


## Syntax

 _expression_ . **Item**( **_Type_** , **_AxisGroup_** )

 _expression_ A variable that represents an **Axes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlAxisType](xlaxistype-enumeration-excel.md)**|The axis type.|
| _AxisGroup_|Optional| **[XlAxisGroup](xlaxisgroup-enumeration-excel.md)**|The axis.|

### Return Value

Axis


## Example

This example sets the title text for the category axis on Chart1.


```vb
With Charts("chart1").Axes.Item(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


## See also


#### Concepts


[Axes Collection](axes-object-excel.md)

