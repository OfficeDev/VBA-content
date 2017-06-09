---
title: PageSetup.LineNumbering Property (Word)
keywords: vbawd10.chm158400630
f1_keywords:
- vbawd10.chm158400630
ms.prod: word
api_name:
- Word.PageSetup.LineNumbering
ms.assetid: acdf1ef4-baaa-aa22-b7a1-81e89d1cebfa
ms.date: 06/08/2017
---


# PageSetup.LineNumbering Property (Word)

Returns or sets a  **[LineNumbering](linenumbering-object-word.md)** object that represents the line numbers for the specified **PageSetup** object.


## Syntax

 _expression_ . **LineNumbering**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

You must be in print layout view to see line numbering.


## Example

This example enables line numbering for the active document.


```vb
ActiveDocument.PageSetup.LineNumbering.Active = True
```

This example enables line numbering for a document named "MyDocument.doc" The starting number is set to one, every fifth line number is shown, and the numbering is continuous throughout all sections in the document.




```vb
set myDoc = Documents("MyDocument.doc") 
With myDoc.PageSetup.LineNumbering 
 .Active = True 
 .StartingNumber = 1 
 .CountBy = 5 
 .RestartMode = wdRestartContinuous 
End With
```

This example sets the line numbering in the active document equal to the line numbering in MyDocument.doc.




```vb
ActiveDocument.PageSetup.LineNumbering = Documents("MyDocument.doc") _ 
 .PageSetup.LineNumbering
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

