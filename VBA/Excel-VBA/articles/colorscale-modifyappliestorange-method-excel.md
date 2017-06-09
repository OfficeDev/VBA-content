---
title: ColorScale.ModifyAppliesToRange Method (Excel)
keywords: vbaxl10.chm806081
f1_keywords:
- vbaxl10.chm806081
ms.prod: excel
api_name:
- Excel.ColorScale.ModifyAppliesToRange
ms.assetid: afa0d0c4-abda-1f16-6b52-a4d330e62dbe
ms.date: 06/08/2017
---


# ColorScale.ModifyAppliesToRange Method (Excel)

Sets the cell range to which this formatting rule applies.


## Syntax

 _expression_ . **ModifyAppliesToRange**( **_Range_** )

 _expression_ A variable that represents a **ColorScale** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range to which this formatting rule will be applied.|

## Remarks

The range must be in the A1 reference style and be entirely contained within the sheet that is the parent of the  **[FormatConditions](formatconditions-object-excel.md)** collection. It can include the range operator (a colon), the intersection operator (a space), or the union operator (a comma). Dollar signs can also be used but they are ignored.

You can also use a local defined name in any part of the range, but the name must be in the language of the macro.


## See also


#### Concepts


[ColorScale Object](colorscale-object-excel.md)

