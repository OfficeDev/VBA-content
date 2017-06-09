---
title: AddReplacement Method
keywords: vbagr10.chm66682
f1_keywords:
- vbagr10.chm66682
ms.prod: excel
api_name:
- Excel.AddReplacement
ms.assetid: 70a6a3f7-e42f-e8b4-d7f8-1ad8f8c66ba7
ms.date: 06/08/2017
---


# AddReplacement Method

Adds an entry to the array of AutoCorrect replacements.

 _expression_. **AddReplacement( _What_**,  **_Replacement_)**

 _expression_ Required. An expression that returns an **AutoCorrect** object.

 **What** Required **String**. The text to be replaced. If this string already exists in the array of AutoCorrect replacements, the existing substitute text is replaced by the new text.
 **Replacement** Required **String**. The replacement text.

## Example

This example substitutes the word "Temp." for the word "Temperature" in the array of AutoCorrect replacements.


```vb
With myChart.Application.AutoCorrect 
 .AddReplacement "Temperature", "Temp." 
End With
```


