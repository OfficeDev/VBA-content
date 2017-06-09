---
title: LineNumbering.Active Property (Word)
keywords: vbawd10.chm158466152
f1_keywords:
- vbawd10.chm158466152
ms.prod: word
api_name:
- Word.LineNumbering.Active
ms.assetid: 31b62e8f-a254-21aa-97bf-d9114f0605a8
ms.date: 06/08/2017
---


# LineNumbering.Active Property (Word)

 **True** if line numbering is active for the specified document, section, or sections. Read/write **Long** .


## Syntax

 _expression_ . **Active**

 _expression_ An expression that returns a **[LineNumbering](linenumbering-object-word.md)** object.


## Example

This example activates line numbering for the first section in the selection.


```vb
Sub CountByFive() 
 With Selection.Sections(1).PageSetup.LineNumbering 
 .Active = True 
 .CountBy = 5 
 .StartingNumber = 1 
 End With 
End Sub
```


## See also


#### Concepts


[LineNumbering Object](linenumbering-object-word.md)

