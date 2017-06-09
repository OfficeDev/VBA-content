---
title: Round Function
keywords: vblr6.chm1009020
f1_keywords:
- vblr6.chm1009020
ms.prod: office
ms.assetid: 897563a8-e66a-1ff1-36b2-da44ae56f48c
ms.date: 06/08/2017
---


# Round Function



 **Description**
Returns a number rounded to a specified number of decimal places.
 **Syntax**
 **Round(**_expression_ [ **,**_numdecimalplaces_ ] **)**
The  **Round** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _expression_|Required. [Numeric expression](vbe-glossary.md) being rounded.|
| _numdecimalplaces_|Optional. Number indicating how many places to the right of the decimal are included in the rounding. If omitted, integers are returned by the  **Round** function.|

 **Note**
This VBA function returns something commonly referred to as bankers rounding. So be careful before using this function. For more predictable results use Worksheet Round functions in Excel VBA:
```
?Round(0.12335,4)
 0,1234
?Round(0.12345,4)
 0,1234
?Round(0.12355,4)
 0,1236
?Round(0.12365,4)
 0,1236

?WorksheetFunction.Round(0.12345,4)
 0,1235
?WorksheetFunction.RoundUp(0.12345,4)
 0,1235
?WorksheetFunction.RoundDown(0.12345,4)
 0,1234

?Round(0.00005,4)
 0
?WorksheetFunction.Round(0.00005,4)
 0,0001
?WorksheetFunction.RoundUp(0.00005,4)
 0,0001
?WorksheetFunction.RoundDown(0.00005,4)
 0
```
