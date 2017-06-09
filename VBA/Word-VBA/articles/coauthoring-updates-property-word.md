---
title: CoAuthoring.Updates Property (Word)
keywords: vbawd10.chm254869510
f1_keywords:
- vbawd10.chm254869510
ms.prod: word
api_name:
- Word.CoAuthoring.Updates
ms.assetid: 89c99cbd-1b97-24b1-f614-d7ade4f383bc
ms.date: 06/08/2017
---


# CoAuthoring.Updates Property (Word)

Returns a  **[CoAuthUpdates](http://msdn.microsoft.com/library/4a164415-0c6c-213b-da94-744e2394d1ef%28Office.15%29.aspx)** collection that represents the most recent updates that were merged into the document. Read-only.


## Syntax

 _expression_ . **Updates**

 _expression_ An expression that returns a **[CoAuthoring](coauthoring-object-word.md)** object.


## Example

The following code example gets the most recent updates that have been merged into the active document.


```vb
Dim allUpdates As CoAuthUpdates 
 
Set allUpdates = ActiveDocument.CoAuthoring.Updates
```


## See also


#### Concepts


[CoAuthoring Object](coauthoring-object-word.md)

