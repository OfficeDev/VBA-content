---
title: Worksheet.BeforeDoubleClick Event (Excel)
keywords: vbaxl10.chm502074
f1_keywords:
- vbaxl10.chm502074
ms.prod: excel
api_name:
- Excel.Worksheet.BeforeDoubleClick
ms.assetid: 36e23bc8-0b49-2e22-bfb0-cfff24a82fda
ms.date: 06/08/2017
---


# Worksheet.BeforeDoubleClick Event (Excel)

Occurs when a worksheet is double-clicked, before the default double-click action.


## Syntax

 _expression_ . **BeforeDoubleClick**( **_Target_** , **_Cancel_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **Range**|The cell nearest to the mouse pointer when the double-click occurs.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the default double-click action isn't performed when the procedure is finished.|

## Remarks

The  **[DoubleClick](application-doubleclick-method-excel.md)** method doesn't cause this event to occur.

This event doesn't occur when the user double-clicks the border of a cell.


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

