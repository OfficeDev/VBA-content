---
title: Worksheet.ClearCircles Method (Excel)
keywords: vbaxl10.chm175141
f1_keywords:
- vbaxl10.chm175141
ms.prod: excel
api_name:
- Excel.Worksheet.ClearCircles
ms.assetid: 74795226-886b-5922-5448-b93355415bd1
ms.date: 06/08/2017
---


# Worksheet.ClearCircles Method (Excel)

Clears circles from invalid entries on the worksheet.


## Syntax

 _expression_ . **ClearCircles**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

Use the  **[CircleInvalid](worksheet-circleinvalid-method-excel.md)** method to circle cells that contain invalid data.


## Example

This example clears circles from invalid entries on worksheet one.


```vb
Worksheets(1).ClearCircles
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

