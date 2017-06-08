---
title: AutoCorrect.TwoInitialCapitals Property (Excel)
keywords: vbaxl10.chm545078
f1_keywords:
- vbaxl10.chm545078
ms.prod: excel
api_name:
- Excel.AutoCorrect.TwoInitialCapitals
ms.assetid: bc24bbfc-fe6d-ca18-c246-49c4c59e9181
ms.date: 06/08/2017
---


# AutoCorrect.TwoInitialCapitals Property (Excel)

 **True** if words that begin with two capital letters are corrected automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **TwoInitialCapitals**

 _expression_ A variable that represents an **AutoCorrect** object.


## Example

This example sets Microsoft Excel to correct words that begin with two capital letters.


```vb
With Application.AutoCorrect 
 .TwoInitialCapitals = True 
 .ReplaceText = True 
End With
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-excel.md)

