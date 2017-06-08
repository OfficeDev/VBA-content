---
title: Scenario.ChangeScenario Method (Excel)
keywords: vbaxl10.chm364073
f1_keywords:
- vbaxl10.chm364073
ms.prod: excel
api_name:
- Excel.Scenario.ChangeScenario
ms.assetid: a982a903-d62c-5289-8192-67f5069a1d11
ms.date: 06/08/2017
---


# Scenario.ChangeScenario Method (Excel)

Changes the scenario to have a new set of changing cells and (optionally) scenario values.


## Syntax

 _expression_ . **ChangeScenario**( **_ChangingCells_** , **_Values_** )

 _expression_ A variable that represents a **Scenario** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ChangingCells_|Required| **Variant**|A  **Range** object that specifies the new set of changing cells for the scenario. The changing cells must be on the same sheet as the scenario.|
| _Values_|Optional| **Variant**|An array that contains the new scenario values for the changing cells. If this argument is omitted, the scenario values are assumed to be the current values in the changing cells.|

### Return Value

Variant


## Remarks

If you specify  _Values_, the array must contain an element for each cell in the  _ChangingCells_ range; otherwise, Microsoft Excel generates an error.


## Example

This example sets the changing cells for scenario one to the range A1:A10 on Sheet1.


```vb
Worksheets("Sheet1").Scenarios(1).ChangeScenario _ 
 Worksheets("Sheet1").Range("A1:A10")
```


## See also


#### Concepts


[Scenario Object](scenario-object-excel.md)

