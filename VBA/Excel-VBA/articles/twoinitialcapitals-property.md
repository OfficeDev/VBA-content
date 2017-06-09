---
title: TwoInitialCapitals Property
keywords: vbagr10.chm5208088
f1_keywords:
- vbagr10.chm5208088
ms.prod: excel
api_name:
- Excel.TwoInitialCapitals
ms.assetid: cf6931c7-11ee-77b0-feb2-e047f7cb58e6
ms.date: 06/08/2017
---


# TwoInitialCapitals Property

 **True** if words that begin with two capital letters are corrected automatically. Read/write **Boolean**.


## Example

This example sets Microsoft Graph to automatically correct words that begin with two capital letters.


```vb
With myChart.Application.AutoCorrect 
 .TwoInitialCapitals = True 
 .ReplaceText = True 
End With
```


