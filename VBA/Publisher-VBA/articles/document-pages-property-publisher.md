---
title: Document.Pages Property (Publisher)
keywords: vbapb10.chm196631
f1_keywords:
- vbapb10.chm196631
ms.prod: publisher
api_name:
- Publisher.Document.Pages
ms.assetid: 2bb3e529-a459-b37c-c9ae-4cc059954a63
ms.date: 06/08/2017
---


# Document.Pages Property (Publisher)

Returns a  **[Pages](pages-object-publisher.md)** collection representing all the pages in the specified publication.


## Syntax

 _expression_. **Pages**

 _expression_A variable that represents a  **Document** object.


## Example

The following example returns the  **Pages** collection of the active publication and reports how many pages there are.


```vb
Dim pgsTemp As Pages 
 
Set pgsTemp = ActiveDocument.Pages 
 
With pgsTemp 
 MsgBox "There are " &; .Count _ 
 &; " page(s) in the active publication." 
End With
```


