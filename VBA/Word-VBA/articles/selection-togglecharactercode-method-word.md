---
title: Selection.ToggleCharacterCode Method (Word)
keywords: vbawd10.chm158663668
f1_keywords:
- vbawd10.chm158663668
ms.prod: word
api_name:
- Word.Selection.ToggleCharacterCode
ms.assetid: e59774bc-cdd5-577b-8175-f988a18c0538
ms.date: 06/08/2017
---


# Selection.ToggleCharacterCode Method (Word)

Switches a selection between a Unicode character and its corresponding hexadecimal value.


## Syntax

 _expression_ . **ToggleCharacterCode**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example enters the hexadecimal value "20ac" at the cursor position and switches that value to its corresponding Unicode character.


```vb
Sub ToggleCharCase() 
 Selection.TypeText Text:="20ac" 
 Selection.ToggleCharacterCode 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

