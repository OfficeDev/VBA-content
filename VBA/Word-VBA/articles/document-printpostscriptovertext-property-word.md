---
title: Document.PrintPostScriptOverText Property (Word)
keywords: vbawd10.chm158007376
f1_keywords:
- vbawd10.chm158007376
ms.prod: word
api_name:
- Word.Document.PrintPostScriptOverText
ms.assetid: 614e3776-c3e7-a4ca-3148-2f285229ecb2
ms.date: 06/08/2017
---


# Document.PrintPostScriptOverText Property (Word)

 **True** if PRINT field instructions (such as PostScript commands) in a document are to be printed on top of text and graphics when a PostScript printer is used. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintPostScriptOverText**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **PrintPostScriptOverText** property controls whether postscript code is printed in a converted Microsoft Word for Macintosh document. If the document contains no PRINT fields, this property has no effect.


## Example

This example sets Word to print PRINT field instructions on top of text and graphics, and then it prints the active document.


```vb
ActiveDocument.PrintPostScriptOverText = True 
ActiveDocument.PrintOut
```

This example returns the current status of the Print PostScript over text check box in the Printing options area on the Print tab in the Options dialog box.




```
currSet = ActiveDocument.PrintPostScriptOverText
```


## See also


#### Concepts


[Document Object](document-object-word.md)

