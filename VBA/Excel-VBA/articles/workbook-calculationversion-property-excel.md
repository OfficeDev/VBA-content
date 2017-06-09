---
title: Workbook.CalculationVersion Property (Excel)
keywords: vbaxl10.chm199192
f1_keywords:
- vbaxl10.chm199192
ms.prod: excel
api_name:
- Excel.Workbook.CalculationVersion
ms.assetid: 09633164-998f-9fa7-f257-da109c369cd7
ms.date: 06/08/2017
---


# Workbook.CalculationVersion Property (Excel)

Returns the information about the version of Excel that the workbook was last fully recalculated by. Read-only  **Long** .


## Syntax

 _expression_ . **CalculationVersion**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

If the workbook was saved in an earlier version of Excel and if the workbook hasn't been fully recalculated, then this property returns 0.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

