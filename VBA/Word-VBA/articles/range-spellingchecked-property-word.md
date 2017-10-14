---
title: Range.SpellingChecked Property (Word)
keywords: vbawd10.chm157155589
f1_keywords:
- vbawd10.chm157155589
ms.prod: word
api_name:
- Word.Range.SpellingChecked
ms.assetid: 5a58fb94-186b-d30c-bef4-d42a295fdeb6
ms.date: 06/08/2017
---


# Range.SpellingChecked Property (Word)

 **True** if spelling has been checked throughout the specified range or document. **False** if all or some of the range or document has not been checked for spelling. Read/write **Boolean** .


## Syntax

 _expression_ . **SpellingChecked**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

To recheck the spelling in a range or document, set the  **SpellingChecked** property to **False** .

To see whether the range or document contains spelling errors, use the  **SpellingErrors** property.


## Example

This example determines whether spelling in section one of the active document has been checked. If spelling has not been checked, the example starts a spelling check.


```vb
Set myRange = ActiveDocument.Sections(1).Range 
isChecked = myRange.SpellingChecked 
If isChecked = False Then 
 myRange.CheckSpelling 
Else 
 MsgBox "Spelling has already been checked in the range." 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

