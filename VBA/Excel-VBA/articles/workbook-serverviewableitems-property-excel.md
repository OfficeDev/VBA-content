---
title: Workbook.ServerViewableItems Property (Excel)
keywords: vbaxl10.chm199245
f1_keywords:
- vbaxl10.chm199245
ms.prod: excel
api_name:
- Excel.Workbook.ServerViewableItems
ms.assetid: 2c10a647-2b2c-0640-9990-109b89444cd2
ms.date: 06/08/2017
---


# Workbook.ServerViewableItems Property (Excel)

Allows a developer to interact with the list of published objects in the workbook that are shown on the server. Read-only.


## Syntax

 _expression_ . **ServerViewableItems**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

This property returns a collection of the items that could be published to Excel Services. It can include  **Tables**,  **PivotTables**,  **Named Ranges**, and  **Charts**. It can also contain  **Sheets** as long as it is not a mixture of **Sheets** and the before mentioned list.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

