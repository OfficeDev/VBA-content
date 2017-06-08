---
title: SpellingOptions.KoreanProcessCompound Property (Excel)
keywords: vbaxl10.chm717082
f1_keywords:
- vbaxl10.chm717082
ms.prod: excel
api_name:
- Excel.SpellingOptions.KoreanProcessCompound
ms.assetid: c6bb9d79-d464-1644-4873-5f3ccf84e487
ms.date: 06/08/2017
---


# SpellingOptions.KoreanProcessCompound Property (Excel)

When set to  **True** , this enables Microsoft Excel to process Korean compound nouns when using the spelling checker. Read/write **Boolean** .


## Syntax

 _expression_ . **KoreanProcessCompound**

 _expression_ A variable that represents a **SpellingOptions** object.


## Example

In this example, Microsoft Excel checks to see if the spell checking option to process Korean compound nouns is on or off and notifies the user accordingly.


```vb
Sub KoreanSpellCheck() 
 
 If Application.SpellingOptions.KoreanProcessCompound = True Then 
 MsgBox "The spell checking feature to process Korean compound nouns is on." 
 Else 
 MsgBox "The spell checking feature to process Korean compound nouns is off." 
 End If 
 
End Sub
```


## See also


#### Concepts


[SpellingOptions Object](spellingoptions-object-excel.md)

