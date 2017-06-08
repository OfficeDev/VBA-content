---
title: ReplaceText Property
keywords: vbagr10.chm5207920
f1_keywords:
- vbagr10.chm5207920
ms.prod: excel
api_name:
- Excel.ReplaceText
ms.assetid: 930c453b-5363-3124-ec06-62359e41ee47
ms.date: 06/08/2017
---


# ReplaceText Property

 **True** if text in the list of AutoCorrect replacements is replaced automatically. Read/write **Boolean**.


## Example

This example turns off automatic text replacement for the chart.


```vb
With myChart.Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = False 
End With
```


