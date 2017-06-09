---
title: Application Property
keywords: vbagr10.chm3076941
f1_keywords:
- vbagr10.chm3076941
ms.prod: excel
api_name:
- Excel.Application
ms.assetid: df183c1c-8db3-e85c-c390-977cf54db7c5
ms.date: 06/08/2017
---


# Application Property

Returns an Application object that represents the Microsoft Graph application. Read-only Application object.

 _expression_. **Application**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example substitutes the word "Temp." for the word "Temperature" in the array of AutoCorrect replacements.


```vb
With myChart.Application.AutoCorrect 
 .AddReplacement "Temperature", "Temp." 
End With
```


