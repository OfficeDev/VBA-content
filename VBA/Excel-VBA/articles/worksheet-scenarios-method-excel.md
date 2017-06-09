---
title: Worksheet.Scenarios Method (Excel)
keywords: vbaxl10.chm175123
f1_keywords:
- vbaxl10.chm175123
ms.prod: excel
api_name:
- Excel.Worksheet.Scenarios
ms.assetid: 52e60b55-9316-4c0b-4cb7-ef4605bd31eb
ms.date: 06/08/2017
---


# Worksheet.Scenarios Method (Excel)

Returns an object that represents either a single scenario (a  **[Scenario](scenario-object-excel.md)** object) or a collection of scenarios (a **[Scenarios](scenarios-object-excel.md)** object) on the worksheet.


## Syntax

 _expression_ . **Scenarios**( **_Index_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the scenario. Use an array to specify more than one scenario.|

### Return Value

Object


## Example

This example sets the comment for the first scenario on Sheet1.


```vb
Worksheets("Sheet1").Scenarios(1).Comment = _ 
 "Worst-case July 1993 sales"
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

