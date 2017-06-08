---
title: AutoCorrect.OtherCorrectionsExceptions Property (Word)
keywords: vbawd10.chm155779089
f1_keywords:
- vbawd10.chm155779089
ms.prod: word
api_name:
- Word.AutoCorrect.OtherCorrectionsExceptions
ms.assetid: 6353059f-1a87-85e6-8783-f7836ea214f1
ms.date: 06/08/2017
---


# AutoCorrect.OtherCorrectionsExceptions Property (Word)

Returns an  **[OtherCorrectionsExceptions](othercorrectionsexceptions-object-word.md)** collection that represents the list of words that Microsoft Word won't correct automatically.


## Syntax

 _expression_ . **OtherCorrectionsExceptions**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Remarks

This list that this property returns corresponds to the list of AutoCorrect exceptions on the  **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box.

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example prompts the user to delete or keep each AutoCorrect exception on the  **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box.


```vb
For Each anEntry In _ 
 AutoCorrect.OtherCorrectionsExceptions 
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

