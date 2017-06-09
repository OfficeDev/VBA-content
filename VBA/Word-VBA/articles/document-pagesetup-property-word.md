---
title: Document.PageSetup Property (Word)
keywords: vbawd10.chm158008397
f1_keywords:
- vbawd10.chm158008397
ms.prod: word
api_name:
- Word.Document.PageSetup
ms.assetid: ddc90b56-f18b-3a30-23d3-24f95d9af8a6
ms.date: 06/08/2017
---


# Document.PageSetup Property (Word)

Returns a  **PageSetup** object that is associated with the specified document.


## Syntax

 _expression_ . **PageSetup**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets the right margin of the active document to 72 points (1 inch).


```vb
ActiveDocument.PageSetup.RightMargin = InchesToPoints(1)
```

This example displays the left margin setting, in inches.




```vb
MsgBox PointsToInches(ActiveDocument.PageSetup.LeftMargin) _ 
 &; " inches"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

