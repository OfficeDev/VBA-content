---
title: SpellingOptions.DictLang Property (Excel)
keywords: vbaxl10.chm717073
f1_keywords:
- vbaxl10.chm717073
ms.prod: excel
api_name:
- Excel.SpellingOptions.DictLang
ms.assetid: 3564b149-5d37-88b4-a0b1-73398e9373c5
ms.date: 06/08/2017
---


# SpellingOptions.DictLang Property (Excel)

Selects the dictionary language used when Microsoft Excel performs spelling checks. Read/write  **Long** .


## Syntax

 _expression_ . **DictLang**

 _expression_ A variable that represents a **SpellingOptions** object.


## Example

This example sets the Excel dictionary to use the English (United States) language.


```vb
Sub LanguageSpellCheck() 
 
 With Application.SpellingOptions 
 .DictLang = 1033 ' United States English language number. 
 .UserDict = "CUSTOM.DIC" 
 End With 
 
End Sub
```


## See also


#### Concepts


[SpellingOptions Object](spellingoptions-object-excel.md)

