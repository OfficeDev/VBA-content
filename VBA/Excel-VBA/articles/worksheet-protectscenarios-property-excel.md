---
title: Worksheet.ProtectScenarios Property (Excel)
keywords: vbaxl10.chm174093
f1_keywords:
- vbaxl10.chm174093
ms.prod: excel
api_name:
- Excel.Worksheet.ProtectScenarios
ms.assetid: 7b0aacea-00f3-7f0a-2be1-693f0efbec88
ms.date: 06/08/2017
---


# Worksheet.ProtectScenarios Property (Excel)

 **True** if the worksheet scenarios are protected. Read-only **Boolean** .


## Syntax

 _expression_ . **ProtectScenarios**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example displays a message box if scenarios are protected on Sheet1.


```vb
If Worksheets("Sheet1").ProtectScenarios Then _ 
 MsgBox "Scenarios are protected on this worksheet."
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

