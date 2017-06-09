---
title: AutoCorrect Object
keywords: vbagr10.chm3077641
f1_keywords:
- vbagr10.chm3077641
ms.prod: excel
api_name:
- Excel.AutoCorrect
ms.assetid: 68fa11da-e28f-53cd-3d50-a1f19d261a02
ms.date: 06/08/2017
---


# AutoCorrect Object

Contains Microsoft Graph AutoCorrect attributes (capitalization of names of days, correction of two initial capital letters, automatic correction list, and so on).


## Using the AutoCorrect Object

Use the  **[AutoCorrect](autocorrect-property.md)** property to return the  **AutoCorrect** object. The following example sets Microsoft Graph to correct words that begin with two initial capital letters.


```vb
With myChart.Application.AutoCorrect 
 .TwoInitialCapitals = True 
 .ReplaceText = True 
End With
```


