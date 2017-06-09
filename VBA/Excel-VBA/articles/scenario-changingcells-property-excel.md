---
title: Scenario.ChangingCells Property (Excel)
keywords: vbaxl10.chm364074
f1_keywords:
- vbaxl10.chm364074
ms.prod: excel
api_name:
- Excel.Scenario.ChangingCells
ms.assetid: 254abee5-0b64-7f68-33e9-28228541ad8f
ms.date: 06/08/2017
---


# Scenario.ChangingCells Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the changing cells for a scenario. Read-only.


## Syntax

 _expression_ . **ChangingCells**

 _expression_ A variable that represents a **Scenario** object.


## Example

This example selects the changing cells for scenario one on Sheet1.


```vb
Worksheets("Sheet1").Activate 
ActiveSheet.Scenarios(1).ChangingCells.Select
```


## See also


#### Concepts


[Scenario Object](scenario-object-excel.md)

