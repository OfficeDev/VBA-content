---
title: Pane.TOCInFrameset Method (Word)
keywords: vbawd10.chm157286507
f1_keywords:
- vbawd10.chm157286507
ms.prod: word
api_name:
- Word.Pane.TOCInFrameset
ms.assetid: fbc96c96-caff-b867-c468-21eec396e014
ms.date: 06/08/2017
---


# Pane.TOCInFrameset Method (Word)

Creates a table of contents based on the specified document and puts it in a new frame on the left side of the frames page.


## Syntax

 _expression_ . **TOCInFrameset**

 _expression_ Required. A variable that represents a **[Pane](pane-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example opens a file named "Proposal.doc", creates a frames page based on the file, and adds a frame (on the left side of the page) containing a table of contents for the file.


```
Documents.Open "C:\Documents\Proposal.doc" 
ActiveDocument.ActiveWindow.ActivePane.NewFrameset 
ActiveDocument.ActiveWindow.ActivePane.TOCInFrameset
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

