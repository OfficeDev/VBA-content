---
title: AutoCorrect.AddReplacement Method (Excel)
keywords: vbaxl10.chm545073
f1_keywords:
- vbaxl10.chm545073
ms.prod: excel
api_name:
- Excel.AutoCorrect.AddReplacement
ms.assetid: 33b83ca0-77b5-00ed-1344-fc5e9a816f74
ms.date: 06/08/2017
---


# AutoCorrect.AddReplacement Method (Excel)

Adds an entry to the array of AutoCorrect replacements.


## Syntax

 _expression_ . **AddReplacement**( **_What_** , **_Replacement_** )

 _expression_ A variable that represents an **AutoCorrect** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _What_|Required| **String**|The text to be replaced. If this string already exists in the array of AutoCorrect replacements, the existing substitute text is replaced by the new text.|
| _Replacement_|Required| **String**|The replacement text.|

### Return Value

Variant


## Example

This example substitutes the word "Temp." for the word "Temperature" in the array of AutoCorrect replacements.


```vb
With Application.AutoCorrect 
 .AddReplacement "Temperature", "Temp." 
End With
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-excel.md)

