---
title: Application.AfterCalculate Event (Excel)
keywords: vbaxl10.chm504103
f1_keywords:
- vbaxl10.chm504103
ms.prod: excel
api_name:
- Excel.Application.AfterCalculate
ms.assetid: ed76a36f-1b52-4464-da44-e64c81fb8d38
ms.date: 06/08/2017
---


# Application.AfterCalculate Event (Excel)

The  **AfterCalculate** event occurs when all pending refresh activity (both synchronous and asynchronous) and all of the resultant calculation activities have been completed.


## Syntax

 _expression_ . **AfterCalculate**

 _expression_ A variable that represents an **Application** object.


## Remarks

This event occurs whenever calculation is completed and there are no outstanding queries. It is mandatory for both conditions to be met before the event occurs. The event can be raised even when there is no sheet data in the workbook, such as whenever calculation finishes for the entire workbook and there are no queries running.

Add-in developers use the  **AfterCalculate** event to know when all the data in the workbook has been fully updated by any queries and/or calculations that may have been in progress.

This event occurs after all  **Worksheet** . **Calculate** , **Chart** . **Calculate** , **AfterRefresh** , and **SheetChange** events. It is the last event to occur after all refresh processing and all calc processing have completed, and it occurs after **Application** . **CalculationState** is set to **xlDone** .


## See also


#### Concepts


[Application Object](application-object-excel.md)

