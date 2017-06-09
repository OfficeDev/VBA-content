---
title: AutoCorrect.FirstLetterExceptions Property (Word)
keywords: vbawd10.chm155779079
f1_keywords:
- vbawd10.chm155779079
ms.prod: word
api_name:
- Word.AutoCorrect.FirstLetterExceptions
ms.assetid: 393a7a13-90eb-ce63-f82a-d1b0a9ae2339
ms.date: 06/08/2017
---


# AutoCorrect.FirstLetterExceptions Property (Word)

Returns a  **[FirstLetterExceptions](firstletterexceptions-object-word.md)** collection that represents the list of abbreviations after which Word won't automatically capitalize the next letter. Read-only.


## Syntax

 _expression_ . **FirstLetterExceptions**

 _expression_ A variable that represents an **[AutoCorrect](autocorrect-object-word.md)** object.


## Remarks

This list corresponds to the list of AutoCorrect exceptions on the  **First Letter** tab in the **AutoCorrect Exceptions** dialog box. For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds "apt." to the list of AutoCorrect First Letter exceptions.


```
AutoCorrect.FirstLetterExceptions.Add "apt."
```

This example deletes the specified AutoCorrect First Letter exception if it exists.




```vb
Dim strException As String 
Dim fleLoop As FirstLetterException 
Dim blnMatch As Boolean 
Dim intConfirm As Integer 
 
strException = _ 
 InputBox("Enter the First Letter exception to delete.") 
blnMatch = False 
 
For Each fleLoop in AutoCorrect.FirstLetterExceptions 
 If fleLoop.Name = strException Then 
 blnMatch = True 
 intConfirm = MsgBox("Are you sure you want to delete " _ 
 &; fleLoop.Name, 4) 
 If intConfirm = vbYes Then 
 fleLoop.Delete 
 End If 
 End If 
Next fleLoop 
 
If blnMatch <> True Then 
 MsgBox "There was no First Letter exception: " _ 
 &; strException 
End If
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

