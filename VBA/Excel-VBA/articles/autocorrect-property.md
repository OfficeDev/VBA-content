---
title: AutoCorrect Property
keywords: vbagr10.chm5207061
f1_keywords:
- vbagr10.chm5207061
ms.prod: excel
api_name:
- Excel.AutoCorrect
ms.assetid: f05a4ff5-4245-ff2e-1082-f48e130d0741
ms.date: 06/08/2017
---


# AutoCorrect Property

Returns an  **[AutoCorrect](autocorrect-object.md)** object that represents the Microsoft Graph AutoCorrect attributes. Read-only.


## Example

This example substitutes the word "Temp." for the word "Temperature" in the array of AutoCorrect replacements.


```vb
With myChart.Application.AutoCorrect 
 .AddReplacement "Temperature", "Temp." 
End With
```


