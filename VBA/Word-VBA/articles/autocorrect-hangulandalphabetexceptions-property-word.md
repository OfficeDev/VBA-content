---
title: AutoCorrect.HangulAndAlphabetExceptions Property (Word)
keywords: vbawd10.chm155779085
f1_keywords:
- vbawd10.chm155779085
ms.prod: word
api_name:
- Word.AutoCorrect.HangulAndAlphabetExceptions
ms.assetid: afb525ff-be41-c260-5210-f6ef930b8b04
ms.date: 06/08/2017
---


# AutoCorrect.HangulAndAlphabetExceptions Property (Word)

Returns a  **[HangulAndAlphabetExceptions](hangulandalphabetexceptions-object-word.md)** collection that represents the list of Hangul and alphabet AutoCorrect exceptions.


## Syntax

 _expression_ . **HangulAndAlphabetExceptions**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Remarks

This list corresponds to the list of Hangul and alphabet AutoCorrect exceptions on the  **Korean** tab in the **AutoCorrect Exceptions** dialog box.

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example prompts the user to delete or keep each hangul and alphabet AutoCorrect exception on the Korean tab in the AutoCorrect Exceptions dialog box.


```vb
For Each anEntry In _ 
 AutoCorrect.HangulAndAlphabetExceptions 
 response = MsgBox("Delete entry: " _ 
 &; anEntry.Name, vbYesNoCancel) 
 If response = vbYes Then 
 anEntry.Delete 
 Else 
 If response = vbCancel Then End 
 End If 
Next anEntry
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

