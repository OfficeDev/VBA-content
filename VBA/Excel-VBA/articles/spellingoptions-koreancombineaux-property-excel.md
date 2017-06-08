---
title: SpellingOptions.KoreanCombineAux Property (Excel)
keywords: vbaxl10.chm717080
f1_keywords:
- vbaxl10.chm717080
ms.prod: excel
api_name:
- Excel.SpellingOptions.KoreanCombineAux
ms.assetid: 9e858f87-e302-2d51-aa9e-383352b534e2
ms.date: 06/08/2017
---


# SpellingOptions.KoreanCombineAux Property (Excel)

When set to  **True** , Microsoft Excel combines Korean auxiliary verbs and adjectives when spelling is checked. Read/write **Boolean** .


## Syntax

 _expression_ . **KoreanCombineAux**

 _expression_ A variable that represents a **SpellingOptions** object.


## Example

In this example, Microsoft Excel checks to see if the option to combine Korean auxiliary verbs and adjectives when checking spelling is on or off and notifies the user accordingly.


```vb
Sub KoreanSpellCheck() 
 
 If Application.SpellingOptions.KoreanCombineAux = True Then 
 MsgBox "The option to combine Korean auxiliary verbs and adjectives while checking spelling is on." 
 Else 
 MsgBox "The option to combine Korean auxiliary verbs and adjectives while checking spelling is off." 
 End If 
 
End Sub
```


## See also


#### Concepts


[SpellingOptions Object](spellingoptions-object-excel.md)

