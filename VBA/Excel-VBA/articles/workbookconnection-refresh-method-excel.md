---
title: WorkbookConnection.Refresh Method (Excel)
keywords: vbaxl10.chm774081
f1_keywords:
- vbaxl10.chm774081
ms.prod: excel
api_name:
- Excel.WorkbookConnection.Refresh
ms.assetid: 5e6f045f-6625-857c-eb55-ac52f70e8fb9
ms.date: 06/08/2017
---


# WorkbookConnection.Refresh Method (Excel)

Refreshes a workbook connection.


## Syntax

 _expression_ . **Refresh**

 _expression_ A variable that represents a **WorkbookConnection** object.


## Remarks

 If the **[DisplayAlerts](application-displayalerts-property-excel.md)** property is **False** , dialog boxes are not displayed and the **Refresh** method fails with the Insufficient Connection Information exception.

A refresh failure for one connection will not have any impact on refresh operations for the other connections.


## See also


#### Concepts


[WorkbookConnection Object](workbookconnection-object-excel.md)

