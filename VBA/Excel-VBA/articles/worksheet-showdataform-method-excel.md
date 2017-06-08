---
title: Worksheet.ShowDataForm Method (Excel)
keywords: vbaxl10.chm175127
f1_keywords:
- vbaxl10.chm175127
ms.prod: excel
api_name:
- Excel.Worksheet.ShowDataForm
ms.assetid: 587a5446-d97e-51d1-d1d9-f5113f8afc0f
ms.date: 06/08/2017
---


# Worksheet.ShowDataForm Method (Excel)

Displays the data form associated with the worksheet.


## Syntax

 _expression_ . **ShowDataForm**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

The macro pauses while you're using the data form. When you close the data form, the macro resumes at the line following the  **ShowDataForm** method.

This method runs the custom data form, if one exists.


## Example

This example displays the data form for Sheet1.


```vb
Worksheets(1).ShowDataForm
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

