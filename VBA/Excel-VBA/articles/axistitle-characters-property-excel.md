---
title: AxisTitle.Characters Property (Excel)
keywords: vbaxl10.chm565079
f1_keywords:
- vbaxl10.chm565079
ms.prod: excel
api_name:
- Excel.AxisTitle.Characters
ms.assetid: 7289fa0d-5e84-6603-5536-5df986b26848
ms.date: 06/08/2017
---


# AxisTitle.Characters Property (Excel)

Returns a  **[Characters](characters-object-excel.md)** object that represents a range of characters within the object text. You can use the **Characters** object to format characters within a text string.


## Syntax

 _expression_ . **Characters**( **_Start_** , **_Length_** )

 _expression_ A variable that represents an **AxisTitle** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the  _Start_ character).|

## Remarks

The  **Characters** object isn't a collection.


## See also


#### Concepts


[AxisTitle Object](axistitle-object-excel.md)

