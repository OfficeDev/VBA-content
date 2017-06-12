---
title: Application.AutoCorrect Property (Excel)
keywords: vbaxl10.chm133081
f1_keywords:
- vbaxl10.chm133081
ms.prod: excel
api_name:
- Excel.Application.AutoCorrect
ms.assetid: e339617e-e086-7324-9240-4db9cfcfcee5
ms.date: 06/08/2017
---


# Application.AutoCorrect Property (Excel)

Returns an  **[AutoCorrect](autocorrect-object-excel.md)** object that represents the Microsoft Excel AutoCorrect attributes. Read-only.


## Syntax

 _expression_ . **AutoCorrect**

 _expression_ A variable that represents an **Application** object.


## Example

This example substitutes the word "Temp." for the word "Temperature" in the array of AutoCorrect replacements.


```vb
With Application.AutoCorrect 
 .AddReplacement "Temperature", "Temp." 
End With
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

