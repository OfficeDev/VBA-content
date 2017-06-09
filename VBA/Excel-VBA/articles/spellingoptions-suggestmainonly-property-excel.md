---
title: SpellingOptions.SuggestMainOnly Property (Excel)
keywords: vbaxl10.chm717076
f1_keywords:
- vbaxl10.chm717076
ms.prod: excel
api_name:
- Excel.SpellingOptions.SuggestMainOnly
ms.assetid: f4a5aa0a-78be-bd98-22e8-b85eac0f4428
ms.date: 06/08/2017
---


# SpellingOptions.SuggestMainOnly Property (Excel)

When set to  **True** , instructs Microsoft Excel to suggest words from only the main dictionary, for using the spelling checker. **False** removes the limits of suggesting words from only the main dictionary, for using the spelling checker. Read/write **Boolean** .


## Syntax

 _expression_ . **SuggestMainOnly**

 _expression_ A variable that represents a **SpellingOptions** object.


## Example

In this example, Microsoft Excel checks the spell checking options for suggesting words only from the main dictionary and reports the status to the user.


```vb
Sub UsingMainDictionary() 
 
 ' Check the setting of suggesting words only from the main dictionary. 
 If Application.SpellingOptions.SuggestMainOnly = True Then 
 MsgBox "Spell checking option suggestions will only come from the main dictionary." 
 Else 
 MsgBox "Spell checking option suggestions are not limited to the main dictionary." 
 End If 
 
End Sub
```


## See also


#### Concepts


[SpellingOptions Object](spellingoptions-object-excel.md)

