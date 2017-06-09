---
title: Document.ConsecutiveHyphensLimit Property (Word)
keywords: vbawd10.chm158007310
f1_keywords:
- vbawd10.chm158007310
ms.prod: word
api_name:
- Word.Document.ConsecutiveHyphensLimit
ms.assetid: 73ff4693-232b-fae3-8077-f6675caede1c
ms.date: 06/08/2017
---


# Document.ConsecutiveHyphensLimit Property (Word)

Returns or sets the maximum number of consecutive lines that can end with hyphens. Read/write.  **Long** .


## Syntax

 _expression_ . **ConsecutiveHyphensLimit**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

If the  **ConsecutiveHyphensLimit** property is set to 0 (zero), any number of consecutive lines can end with hyphens.


## Example

This example enables automatic hyphenation for MyReport.doc and limits the number of consecutive lines that can end with hyphens to two.


```vb
With Documents("MyReport.doc") 
 .AutoHyphenation = True 
 .ConsecutiveHyphensLimit = 2 
End With
```

This example sets no limit on the number of consecutive lines that can end with hyphens.




```vb
ActiveDocument.ConsecutiveHyphensLimit = 0
```


## See also


#### Concepts


[Document Object](document-object-word.md)

